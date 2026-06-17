/* global Webex */

// --- Webex OAuth app credentials ---
const CLIENT_ID = 'C7c3f1143a552d88d40b2afff87600c366c830850231597fb6c1c1e28a5110a4f';
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

function appendTranscript(text) {
  const line = document.createElement('div');
  line.innerText = text;
  transcriptEl.appendChild(line);
  transcriptEl.scrollTop = transcriptEl.scrollHeight;
}

function initWebex() {
  const redirectUri = `${location.protocol}//${location.host}${location.pathname}`;

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
      appendTranscript(payload.transcription);
    }
  });

  setStatus('Joining meeting…');
  await meeting.join();

  setStatus('Starting transcription…');
  await meeting.receiveTranscription();

  setStatus('Transcribing live…');
}

initWebex();
