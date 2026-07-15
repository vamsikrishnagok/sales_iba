/* global Webex */

// --- Webex OAuth app credentials ---
const CLIENT_ID = 'Cd7ffc2137ec3ab40ea41a5e20667f9b911baf16fa9697da8ac5b7dcc64d47a2b';
const SCOPE = 'spark:all spark:kms';

// Base URL of the FastAPI backend (see backend/main.py).
// Use localhost while developing; swap to your deployed URL in production.
const API_BASE = 'http://localhost:8000';

const loginBtn = document.querySelector('#login-btn');
const statusEl = document.querySelector('#status');
const transcriptEl = document.querySelector('#transcript');
const aiAnswersEl = document.querySelector('#ai-answers');

// Remember questions we've already sent so we don't ask the LLM twice.
const askedQuestions = new Set();

let webex;
let meeting;

function setStatus(text) {
  statusEl.innerText = text;
  console.log('[status]', text);
}

// Ping the backend so we know it is reachable.
async function checkBackend() {
  try {
    const res = await fetch(`${API_BASE}/health`);
    const data = await res.json();
    console.log('[backend] healthy:', data);
  } catch (err) {
    console.warn('[backend] not reachable:', err.message);
  }
}

// Send one finalized transcript line to the backend.
async function sendTranscriptToBackend(speaker, text, transcriptId) {
  try {
    await fetch(`${API_BASE}/transcripts`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        speaker,
        text,
        is_final: true,
        transcript_id: transcriptId,
      }),
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
    const res = await fetch(`${API_BASE}/ask`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question, speaker }),
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
    sendTranscriptToBackend(speaker, text, id);
  }
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

// Tell the backend the meeting is over so it saves the full transcript to a file.
let meetingEnded = false;
async function saveTranscriptOnMeetingEnd() {
  if (meetingEnded) return;
  meetingEnded = true;
  setStatus('Meeting ended — saving transcript…');
  try {
    const res = await fetch(`${API_BASE}/meeting/end`, { method: 'POST' });
    const data = await res.json();
    console.log('[backend] transcript saved:', data);
    setStatus(
      data.saved
        ? `Transcript saved (${data.lines} lines).`
        : 'Meeting ended — no transcript to save.'
    );
  } catch (err) {
    console.warn('[backend] failed to save transcript:', err.message);
    setStatus(`Meeting ended — failed to save transcript: ${err.message}`);
  }
}

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

  setStatus('Joining meeting…');
  await meeting.join();

  setStatus('Starting transcription…');
  await meeting.receiveTranscription();

  setStatus('Transcribing live…');
}

initWebex();
