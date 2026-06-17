/* global Webex */

// --- Webex OAuth app credentials ---
const CLIENT_ID = 'Cd7ffc2137ec3ab40ea41a5e20667f9b911baf16fa9697da8ac5b7dcc64d47a2b';
const SCOPE = 'spark:all spark:kms';

const loginBtn = document.querySelector('#login-btn');
const statusEl = document.querySelector('#status');
const transcriptEl = document.querySelector('#transcript');

let webex;
let meeting;

function setStatus(text) {
  statusEl.innerText = text;
  console.log('[status]', text);
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

  // Once finalized, stop tracking so the next utterance starts a fresh line.
  if (isFinal) {
    transcriptLines.delete(id);
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

async function joinAndTranscribe() {
  // Transcript chunks arrive through this event in the v3 SDK.
  meeting.on('meeting:receiveTranscription:started', (payload) => {
    if (payload && payload.transcription) {
      renderTranscript(payload);
    }
  });

  setStatus('Joining meeting…');
  await meeting.join();

  setStatus('Starting transcription…');
  await meeting.receiveTranscription();

  setStatus('Transcribing live…');
}

initWebex();
