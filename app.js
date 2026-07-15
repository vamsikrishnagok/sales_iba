/* global Webex */

// --- Webex OAuth app credentials ---
const CLIENT_ID = 'Cd7ffc2137ec3ab40ea41a5e20667f9b911baf16fa9697da8ac5b7dcc64d47a2b';
const SCOPE = 'spark:all spark:kms';

// Base URL of the FastAPI backend (data injection service).
// Use localhost while developing; swap to your deployed URL in production.
const API_BASE = 'http://localhost:8000';

const loginBtn = document.querySelector('#login-btn');
const saveBtn = document.querySelector('#save-btn');
const opportunityIdEl = document.querySelector('#opportunity-id');
const statusEl = document.querySelector('#status');
const transcriptEl = document.querySelector('#transcript');
const aiAnswersEl = document.querySelector('#ai-answers');

// Portal auth elements.
const authSection = document.querySelector('#auth-section');
const appSection = document.querySelector('#app-section');
const authUsernameEl = document.querySelector('#auth-username');
const authPasswordEl = document.querySelector('#auth-password');
const authBtn = document.querySelector('#auth-btn');
const authStatusEl = document.querySelector('#auth-status');
const authUserEl = document.querySelector('#auth-user');
const logoutBtn = document.querySelector('#logout-btn');

const TOKEN_KEY = 'iba_portal_token';
const USER_KEY = 'iba_portal_user';

function getToken() {
  return localStorage.getItem(TOKEN_KEY) || '';
}

function setSession(token, user) {
  localStorage.setItem(TOKEN_KEY, token);
  if (user) localStorage.setItem(USER_KEY, JSON.stringify(user));
}

function clearSession() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(USER_KEY);
}

function getStoredUser() {
  try {
    return JSON.parse(localStorage.getItem(USER_KEY) || 'null');
  } catch {
    return null;
  }
}

// Authenticated fetch against the data injection service. Injects the bearer
// token and JSON content type; on 401 it clears the session and shows login.
async function apiFetch(path, { method = 'GET', body, auth = true } = {}) {
  const headers = { 'Content-Type': 'application/json' };
  if (auth) {
    const token = getToken();
    if (token) headers.Authorization = `Bearer ${token}`;
  }
  const res = await fetch(`${API_BASE}${path}`, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });
  if (res.status === 401 && auth) {
    clearSession();
    showLogin('Session expired — please sign in again.');
    throw new Error('Not authenticated');
  }
  return res;
}

// Remember questions we've already sent so we don't ask the LLM twice.
const askedQuestions = new Set();

// Every finalized transcript line, kept client-side so the Save button can
// persist the complete transcript exactly as shown, regardless of backend state.
const finalizedLines = [];

let webex;
let meeting;

function setStatus(text) {
  statusEl.innerText = text;
  console.log('[status]', text);
}

// Ping the backend so we know it is reachable.
async function checkBackend() {
  try {
    const res = await apiFetch('/health', { auth: false });
    const data = await res.json();
    console.log('[backend] healthy:', data);
  } catch (err) {
    console.warn('[backend] not reachable:', err.message);
  }
}

// Send one finalized transcript line to the backend live session.
async function sendTranscriptToBackend(speaker, text, transcriptId) {
  try {
    await apiFetch('/transcripts/line', {
      method: 'POST',
      body: {
        speaker,
        text,
        is_final: true,
        transcript_id: transcriptId,
        opportunity_id: (opportunityIdEl?.value || '').trim() || null,
      },
    });
  } catch (err) {
    console.warn('[backend] failed to save transcript:', err.message);
  }
}

// Ask the LLM a question and render the answer in the AI Answers section.
async function askLLM(question, speaker) {
  // Clear the placeholder on first use.
  if (askedQuestions.size === 0) {
    aiAnswersEl.innerHTML = '';
  }
  askedQuestions.add(question);

  const block = document.createElement('div');
  block.style.marginBottom = '12px';

  const qEl = document.createElement('div');
  const qLabel = document.createElement('strong');
  qLabel.innerText = `Q (${speaker || 'Unknown'}): `;
  qEl.appendChild(qLabel);
  qEl.appendChild(document.createTextNode(question));

  const aEl = document.createElement('div');
  aEl.style.color = '#0b5394';
  aEl.innerText = 'Thinking…';

  block.appendChild(qEl);
  block.appendChild(aEl);
  aiAnswersEl.appendChild(block);
  aiAnswersEl.scrollTop = aiAnswersEl.scrollHeight;

  try {
    const res = await apiFetch('/transcripts/ask', {
      method: 'POST',
      body: { question, speaker },
    });
    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }
    const data = await res.json();
    aEl.innerText = `A: ${data.answer}`;
  } catch (err) {
    aEl.style.color = '#cc0000';
    aEl.innerText = `A: (failed to get answer — ${err.message})`;
  }
  aiAnswersEl.scrollTop = aiAnswersEl.scrollHeight;
}

// Tracks the live DOM line for each in-progress utterance (keyed by transcript id),
// so interim updates replace the same line instead of stacking new ones.
const transcriptLines = new Map();

// Resolve a speaker's display name from the meeting's member list using the
// personID the transcription payload carries. Falls back gracefully.
function resolveSpeakerName(personID) {
  if (!personID || !meeting) return 'Unknown';
  try {
    const members = meeting.members?.membersCollection?.members || {};
    const list = Array.isArray(members) ? members : Object.values(members);
    const match = list.find(
      (m) => m?.id === personID || m?.participant?.person?.id === personID
    );
    return (
      match?.name ||
      match?.participant?.person?.name ||
      'Unknown'
    );
  } catch {
    return 'Unknown';
  }
}

function renderTranscript(payload) {
  if (!payload) return;

  const text = (payload.transcription || '').trim();
  if (!text) return;

  const id = payload.id || `anon-${Date.now()}`;
  const isFinal = payload.type === 'transcript_final_result';
  const speaker = resolveSpeakerName(payload.personID);

  let line = transcriptLines.get(id);
  if (!line) {
    line = document.createElement('div');
    transcriptEl.appendChild(line);
    transcriptLines.set(id, line);
  }

  line.innerHTML = '';
  const nameEl = document.createElement('strong');
  nameEl.innerText = `${speaker}: `;
  line.appendChild(nameEl);
  line.appendChild(document.createTextNode(text));
  transcriptEl.scrollTop = transcriptEl.scrollHeight;

  // Once finalized, stop tracking so the next utterance starts a fresh line,
  // and forward the completed line to the backend.
  if (isFinal) {
    transcriptLines.delete(id);
    finalizedLines.push({ speaker, text, transcript_id: id });
    if (saveBtn) saveBtn.disabled = false;
    sendTranscriptToBackend(speaker, text, id);
  }
}

// Send the full transcript currently held client-side to the backend to save.
async function saveFullTranscript() {
  if (!finalizedLines.length) {
    setStatus('Nothing to save yet.');
    return;
  }
  const opportunityId = (opportunityIdEl?.value || '').trim();
  if (!opportunityId) {
    setStatus('Please enter an Opportunity ID before saving.');
    opportunityIdEl?.focus();
    return;
  }
  if (saveBtn) saveBtn.disabled = true;
  setStatus('Saving transcript…');
  try {
    const res = await apiFetch('/transcripts/save', {
      method: 'POST',
      body: { opportunity_id: opportunityId, lines: finalizedLines },
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    console.log('[backend] transcript saved:', data);
    setStatus(`Transcript saved (${data.lines} lines) for opportunity ${data.opportunity_id}.`);
  } catch (err) {
    console.warn('[backend] failed to save transcript:', err.message);
    setStatus(`Failed to save transcript: ${err.message}`);
  } finally {
    if (saveBtn) saveBtn.disabled = false;
  }
}

if (saveBtn) {
  saveBtn.addEventListener('click', saveFullTranscript);
}

// Must EXACTLY match a Redirect URI registered in your Webex Integration
// (developer.webex.com -> My Webex Apps -> your Integration -> Redirect URI(s)).
const REDIRECT_URI = 'https://vamsikrishnagok.github.io/sales_iba/';

function initWebex() {
  const redirectUri = REDIRECT_URI;

  webex = window.webex = Webex.init({
    config: {
      appName: 'meeting-transcription',
      meetings: { reconnection: { enabled: true } },
      credentials: {
        client_id: CLIENT_ID,
        redirect_uri: redirectUri,
        scope: SCOPE,
      },
    },
  });

  checkBackend();

  webex.once('ready', () => {
    if (webex.canAuthorize) {
      // Already authenticated (e.g. returning from the OAuth redirect) -> automate everything.
      runMeetingFlow();
    } else {
      loginBtn.disabled = false;
      setStatus('Please log in.');
    }
  });
}

loginBtn.addEventListener('click', () => {
  loginBtn.disabled = true;
  setStatus('Redirecting to login…');
  webex.authorization.initiateLogin();
});

// register -> sync -> join current meeting -> start transcription, fully automated.
async function runMeetingFlow() {
  try {
    loginBtn.style.display = 'none';

    setStatus('Registering…');
    if (!webex.meetings.registered) {
      await webex.meetings.register();
    }

    setStatus('Looking for current meeting…');
    await webex.meetings.syncMeetings();

    meeting = getCurrentMeeting();
    if (meeting) {
      await joinAndTranscribe();
    } else {
      setStatus('Waiting for a meeting to start…');
      webex.meetings.on('meeting:added', async ({ meeting: added }) => {
        if (!meeting) {
          meeting = added;
          await joinAndTranscribe();
        }
      });
    }
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err.message || err}`);
  }
}

function getCurrentMeeting() {
  const all = webex.meetings.getAllMeetings();
  const ids = Object.keys(all);
  return ids.length ? all[ids[0]] : null;
}

// Tell the backend the meeting is over so it saves the full transcript.
let meetingEnded = false;
async function saveTranscriptOnMeetingEnd() {
  if (meetingEnded) return;
  meetingEnded = true;
  const opportunityId = (opportunityIdEl?.value || '').trim();
  if (!finalizedLines.length) {
    setStatus('Meeting ended — no transcript to save.');
    return;
  }
  if (!opportunityId) {
    setStatus('Meeting ended — enter an Opportunity ID and click Save to persist.');
    return;
  }
  setStatus('Meeting ended — saving transcript…');
  try {
    const res = await apiFetch('/transcripts/save', {
      method: 'POST',
      body: { opportunity_id: opportunityId, lines: finalizedLines },
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    console.log('[backend] transcript saved:', data);
    setStatus(`Transcript saved (${data.lines} lines) for opportunity ${data.opportunity_id}.`);
  } catch (err) {
    console.warn('[backend] failed to save transcript:', err.message);
    setStatus(`Meeting ended — failed to save transcript: ${err.message}`);
  }
}

// Last-resort save when the tab/window is closed. `fetch` with keepalive can
// carry the Authorization header (unlike sendBeacon) and survives unload.
function saveTranscriptOnUnload() {
  if (meetingEnded) return;
  const opportunityId = (opportunityIdEl?.value || '').trim();
  const token = getToken();
  if (!finalizedLines.length || !opportunityId || !token) return;
  try {
    fetch(`${API_BASE}/transcripts/save`, {
      method: 'POST',
      keepalive: true,
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({ opportunity_id: opportunityId, lines: finalizedLines }),
    });
  } catch (err) {
    console.warn('[backend] keepalive save failed:', err.message);
  }
}
window.addEventListener('pagehide', saveTranscriptOnUnload);
window.addEventListener('beforeunload', saveTranscriptOnUnload);
async function joinAndTranscribe() {
  // Transcript chunks arrive through this event in the v3 SDK.
  meeting.on('meeting:receiveTranscription:started', (payload) => {
    if (payload && payload.transcription) {
      renderTranscript(payload);
    }
  });

  // When the meeting ends or we leave, persist the full transcript.
  meeting.on('meeting:self:left', saveTranscriptOnMeetingEnd);
  meeting.on('meeting:ended', saveTranscriptOnMeetingEnd);
  webex.meetings.on('meeting:removed', ({ meetingId }) => {
    if (!meeting || meetingId === meeting.id) {
      saveTranscriptOnMeetingEnd();
    }
  });

  setStatus('Joining meeting…');  await meeting.join();

  setStatus('Starting transcription…');
  await meeting.receiveTranscription();

  setStatus('Transcribing live…');
}

// ---------------------------------------------------------------------------
// Portal auth bootstrap
// ---------------------------------------------------------------------------

let webexInitialized = false;

function showLogin(message) {
  if (authSection) authSection.hidden = false;
  if (appSection) appSection.hidden = true;
  if (authStatusEl) authStatusEl.innerText = message || '';
  if (authBtn) authBtn.disabled = false;
}

function showApp() {
  const user = getStoredUser();
  if (authUserEl) authUserEl.innerText = user?.username || 'user';
  if (authSection) authSection.hidden = true;
  if (appSection) appSection.hidden = false;
  // Start the Webex flow once, now that we have a portal token.
  if (!webexInitialized) {
    webexInitialized = true;
    initWebex();
  }
}

async function signIn() {
  const username = (authUsernameEl?.value || '').trim();
  const password = authPasswordEl?.value || '';
  if (!username || !password) {
    if (authStatusEl) authStatusEl.innerText = 'Enter your username and password.';
    return;
  }
  if (authBtn) authBtn.disabled = true;
  if (authStatusEl) authStatusEl.innerText = 'Signing in…';
  try {
    const res = await apiFetch('/auth/login', {
      method: 'POST',
      auth: false,
      body: { username, password },
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      throw new Error(data.detail || `Sign in failed (${res.status})`);
    }
    setSession(data.token, data.user);
    if (authPasswordEl) authPasswordEl.value = '';
    showApp();
  } catch (err) {
    if (authStatusEl) authStatusEl.innerText = err.message;
    if (authBtn) authBtn.disabled = false;
  }
}

function signOut() {
  // Best-effort server-side logout, then clear local state.
  apiFetch('/auth/logout', { method: 'POST' }).catch(() => {});
  clearSession();
  // Reload so the Webex SDK/meeting state is fully reset.
  window.location.reload();
}

if (authBtn) authBtn.addEventListener('click', signIn);
if (authPasswordEl) {
  authPasswordEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') signIn();
  });
}
if (logoutBtn) logoutBtn.addEventListener('click', signOut);

// On load: resume an existing session or prompt for sign-in.
if (getToken()) {
  showApp();
} else {
  showLogin('Please sign in to continue.');
}

