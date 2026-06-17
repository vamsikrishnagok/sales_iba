/* global Webex */
/*
 * Webex Embedded App: Live Transcript Forwarder
 *
 * - Initializes the Webex Embedded App Framework SDK so the page can run
 *   inside Webex (meetings or spaces).
 * - Uses the Webex JavaScript SDK to join the active meeting with
 *   `receiveTranscription: true` and listens for transcription events.
 * - Forwards each transcription payload to a user-supplied FastAPI endpoint
 *   via HTTP POST.
 */

(function () {
  "use strict";

  const els = {
    backendUrl: document.getElementById("backendUrl"),
    accessToken: document.getElementById("accessToken"),
    meetingDestination: document.getElementById("meetingDestination"),
    btnRegister: document.getElementById("btnRegister"),
    btnJoin: document.getElementById("btnJoin"),
    btnLeave: document.getElementById("btnLeave"),
    status: document.getElementById("status"),
    embeddedCtx: document.getElementById("embeddedCtx"),
    transcriptBox: document.getElementById("transcriptBox"),
    log: document.getElementById("log"),
  };

  let webexSdk = null;
  let activeMeeting = null;
  let embeddedApp = null;

  // ---------- helpers ----------

  function log(msg, obj) {
    const ts = new Date().toISOString().split("T")[1].replace("Z", "");
    const line = obj !== undefined ? `[${ts}] ${msg} ${safeJson(obj)}` : `[${ts}] ${msg}`;
    els.log.textContent = (line + "\n" + els.log.textContent).slice(0, 8000);
    console.log(line);
  }

  function safeJson(obj) {
    try {
      return JSON.stringify(obj);
    } catch (e) {
      return String(obj);
    }
  }

  function setStatus(text, kind) {
    els.status.textContent = text;
    els.status.className = "status status-" + (kind || "idle");
  }

  function persistInputs() {
    try {
      localStorage.setItem("backendUrl", els.backendUrl.value.trim());
      localStorage.setItem("meetingDestination", els.meetingDestination.value.trim());
    } catch (_) {
      /* ignore */
    }
  }

  function restoreInputs() {
    try {
      els.backendUrl.value = localStorage.getItem("backendUrl") || "";
      els.meetingDestination.value = localStorage.getItem("meetingDestination") || "";
    } catch (_) {
      /* ignore */
    }
  }

  // ---------- Webex Embedded App Framework ----------

  function initEmbeddedFramework() {
    if (typeof window.Webex === "undefined" || !window.Webex.Application) {
      els.embeddedCtx.textContent =
        "Webex Embedded App SDK not detected. " +
        "This page only exposes the in-Webex features when loaded inside a Webex client.";
      return;
    }

    try {
      embeddedApp = new window.Webex.Application();
      embeddedApp
        .onReady()
        .then(() => {
          log("Embedded App framework ready");
          const ctx = {
            isShared: embeddedApp.isShared,
            theme: embeddedApp.theme,
            language: embeddedApp.language,
            about: embeddedApp.about,
            applicationInfo: embeddedApp.applicationInfo,
          };
          els.embeddedCtx.textContent = JSON.stringify(ctx, null, 2);
        })
        .catch((err) => {
          log("Embedded App onReady() failed", err && err.message);
          els.embeddedCtx.textContent =
            "Embedded App onReady() failed (likely because this is not loaded inside Webex).";
        });
    } catch (err) {
      log("Embedded App init error", err && err.message);
    }
  }

  // ---------- Webex JS SDK ----------

  async function registerSdk() {
    const token = els.accessToken.value.trim();
    if (!token) {
      alert("Please paste a Webex personal access token first.");
      return;
    }
    if (typeof window.Webex === "undefined" || !window.Webex.init) {
      alert(
        "Webex JS SDK failed to load. Check your network or the CDN URL in index.html."
      );
      return;
    }

    setStatus("Registering...", "pending");
    try {
      webexSdk = window.Webex.init({
        config: {
          credentials: { access_token: token },
        },
      });

      await webexSdk.meetings.register();
      await webexSdk.meetings.syncMeetings();

      log("Webex SDK registered & meetings synced");
      setStatus("Registered", "ok");
      els.btnJoin.disabled = false;
    } catch (err) {
      log("registerSdk failed", err && (err.message || err));
      setStatus("Register failed", "err");
    }
  }

  async function joinMeeting() {
    if (!webexSdk) {
      alert("Register the SDK first.");
      return;
    }
    const destination = els.meetingDestination.value.trim();
    if (!destination) {
      alert("Enter the meeting destination (SIP, meeting number, or link).");
      return;
    }

    persistInputs();
    setStatus("Joining...", "pending");

    try {
      const meeting = await webexSdk.meetings.create(destination);
      activeMeeting = meeting;
      bindMeetingEvents(meeting);

      await meeting.join({ receiveTranscription: true });

      log("Joined meeting; awaiting transcription...");
      setStatus("In meeting", "ok");
      els.btnJoin.disabled = true;
      els.btnLeave.disabled = false;
    } catch (err) {
      log("joinMeeting failed", err && (err.message || err));
      setStatus("Join failed", "err");
    }
  }

  async function leaveMeeting() {
    if (!activeMeeting) return;
    setStatus("Leaving...", "pending");
    try {
      await activeMeeting.leave();
      log("Left meeting");
    } catch (err) {
      log("leaveMeeting error", err && (err.message || err));
    } finally {
      activeMeeting = null;
      els.btnJoin.disabled = false;
      els.btnLeave.disabled = true;
      setStatus("Idle", "idle");
    }
  }

  function bindMeetingEvents(meeting) {
    meeting.on("meeting:receiveTranscription:started", (payload) => {
      handleTranscription(payload);
    });

    meeting.on("meeting:receiveTranscription:stopped", () => {
      log("Transcription stopped by Webex");
    });

    meeting.on("error", (err) => {
      log("Meeting error", err && (err.message || err));
    });

    meeting.on("meeting:self:left", () => {
      log("Self left event");
      setStatus("Idle", "idle");
      els.btnJoin.disabled = false;
      els.btnLeave.disabled = true;
      activeMeeting = null;
    });
  }

  // ---------- Transcription handling ----------

  function handleTranscription(payload) {
    renderTranscript(payload);
    forwardToBackend(payload);
  }

  function renderTranscript(payload) {
    const line = document.createElement("div");
    line.className = "line" + (payload.type === "interim" ? " interim" : "");

    const speaker = document.createElement("span");
    speaker.className = "speaker";
    speaker.textContent = payload.personID ? `User ${shortId(payload.personID)}` : "Speaker";

    const ts = document.createElement("span");
    ts.className = "ts";
    ts.textContent = payload.timestamp || new Date().toISOString();

    const text = document.createElement("div");
    text.textContent = payload.transcription || "";

    line.appendChild(speaker);
    line.appendChild(ts);
    line.appendChild(text);

    els.transcriptBox.appendChild(line);
    els.transcriptBox.scrollTop = els.transcriptBox.scrollHeight;
  }

  function shortId(id) {
    return String(id).slice(0, 8);
  }

  async function forwardToBackend(payload) {
    const url = els.backendUrl.value.trim();
    if (!url) return;

    const body = {
      id: payload.id,
      personId: payload.personID,
      transcription: payload.transcription,
      timestamp: payload.timestamp || new Date().toISOString(),
      type: payload.type || "final",
      meetingId: activeMeeting && activeMeeting.id,
    };

    try {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
        mode: "cors",
      });
      if (!res.ok) {
        log("Backend non-2xx", res.status);
      }
    } catch (err) {
      log("Backend POST failed", err && err.message);
    }
  }

  // ---------- wire up ----------

  function init() {
    restoreInputs();
    els.backendUrl.addEventListener("change", persistInputs);
    els.meetingDestination.addEventListener("change", persistInputs);

    els.btnRegister.addEventListener("click", registerSdk);
    els.btnJoin.addEventListener("click", joinMeeting);
    els.btnLeave.addEventListener("click", leaveMeeting);

    initEmbeddedFramework();
    log("App initialized");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
