/* global Webex */
/*
 * Webex Embedded App: Live Transcript Viewer
 *
 * - Initializes the Webex Embedded App Framework SDK so the page can run
 *   inside Webex (meetings or spaces).
 * - Uses the Webex JavaScript SDK to join the active meeting with
 *   `receiveTranscription: true` and listens for transcription events.
 * - Displays transcription payloads in the UI (forwarding disabled).
 */

(function () {
  "use strict";

  const els = {
    clientId: document.getElementById("clientId"),
    redirectUri: document.getElementById("redirectUri"),
    accessToken: document.getElementById("accessToken"),
    meetingDestination: document.getElementById("meetingDestination"),
    btnLogin: document.getElementById("btnLogin"),
    btnLogout: document.getElementById("btnLogout"),
    btnRegister: document.getElementById("btnRegister"),
    btnJoin: document.getElementById("btnJoin"),
    btnLeave: document.getElementById("btnLeave"),
    status: document.getElementById("status"),
    authStatus: document.getElementById("authStatus"),
    embeddedCtx: document.getElementById("embeddedCtx"),
    transcriptBox: document.getElementById("transcriptBox"),
    log: document.getElementById("log"),
  };

  // Scopes required for the meetings SDK + real-time transcription.
  const OAUTH_SCOPE = "spark:all spark:kms";

  let webexSdk = null;
  let activeMeeting = null;
  let embeddedApp = null;
  let detectedMeetingInfo = null; // auto-detected from embedded context

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

  function normalizeToken(rawToken) {
    const token = String(rawToken || "").trim();
    if (!token) return "";
    // Users often paste "Bearer <token>" from curl snippets.
    return token.replace(/^Bearer\s+/i, "").trim();
  }

  async function validateAccessToken(token) {
    const res = await fetch("https://webexapis.com/v1/people/me", {
      method: "GET",
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`Token check failed (${res.status}): ${body.slice(0, 200)}`);
    }
    return res.json();
  }

  function persistInputs() {
    try {
      localStorage.setItem("meetingDestination", els.meetingDestination.value.trim());
      localStorage.setItem("webexClientId", els.clientId.value.trim());
    } catch (_) {
      /* ignore */
    }
  }

  function restoreInputs() {
    try {
      els.meetingDestination.value = localStorage.getItem("meetingDestination") || "";
      els.clientId.value = localStorage.getItem("webexClientId") || "";
    } catch (_) {
      /* ignore */
    }
  }

  // The redirect URI must EXACTLY match one registered on the Integration.
  // Strip any OAuth hash/query the SDK may have appended on return.
  function getRedirectUri() {
    return window.location.origin + window.location.pathname;
  }

  // ---------- Webex OAuth login ----------

  function setAuthStatus(text, kind) {
    if (!els.authStatus) return;
    els.authStatus.textContent = text;
    els.authStatus.className = "status status-" + (kind || "idle");
  }

  // Initializes the JS SDK in OAuth mode. On return from the Webex login
  // redirect, the browser auth plugin automatically parses the access token
  // from the URL hash, yielding a downscopable token (unlike a personal token).
  function initOAuthSdk(clientId) {
    return window.Webex.init({
      config: {
        credentials: {
          client_id: clientId,
          redirect_uri: getRedirectUri(),
          scope: OAUTH_SCOPE,
        },
      },
    });
  }

  async function loginWithWebex() {
    const clientId = els.clientId.value.trim();
    if (!clientId) {
      alert("Enter your Webex Integration Client ID first.");
      return;
    }
    if (typeof window.Webex === "undefined" || !window.Webex.init) {
      alert("Webex JS SDK failed to load. Check your network or the CDN URL.");
      return;
    }
    persistInputs();
    setAuthStatus("Redirecting to Webex...", "pending");
    try {
      webexSdk = initOAuthSdk(clientId);
      log("Initiating Webex OAuth login", { redirectUri: getRedirectUri() });
      await webexSdk.authorization.initiateLogin();
    } catch (err) {
      log("loginWithWebex failed", err && (err.message || err));
      setAuthStatus("Login failed", "err");
    }
  }

  // Called on page load. If a client_id is stored, init the SDK so the auth
  // plugin can parse a token returned in the URL hash after a login redirect.
  async function restoreOAuthSession() {
    const clientId = els.clientId.value.trim();
    if (!clientId || typeof window.Webex === "undefined" || !window.Webex.init) {
      return;
    }
    try {
      webexSdk = initOAuthSdk(clientId);
      // Give the browser auth plugin a moment to parse the hash and load.
      await new Promise((resolve) => {
        if (webexSdk.canAuthorize) return resolve();
        const done = () => resolve();
        if (typeof webexSdk.once === "function") webexSdk.once("ready", done);
        setTimeout(done, 4000);
      });
      if (webexSdk.canAuthorize) {
        log("Restored Webex OAuth session", "canAuthorize=true");
        onSignedIn();
      } else {
        setAuthStatus("Not signed in", "idle");
      }
    } catch (err) {
      log("restoreOAuthSession error", err && (err.message || err));
    }
  }

  function onSignedIn() {
    setAuthStatus("Signed in", "ok");
    els.btnLogin.disabled = true;
    els.btnLogout.disabled = false;
    els.btnRegister.disabled = false;
  }

  async function logout() {
    try {
      if (webexSdk && webexSdk.logout) {
        await webexSdk.logout();
      }
    } catch (err) {
      log("logout error", err && (err.message || err));
    } finally {
      webexSdk = null;
      setAuthStatus("Not signed in", "idle");
      els.btnLogin.disabled = false;
      els.btnLogout.disabled = true;
      els.btnJoin.disabled = true;
      setStatus("Idle", "idle");
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

          // Try to detect the current meeting context
          tryDetectMeeting(embeddedApp);
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

  function tryDetectMeeting(app) {
    // The Embedded App Framework may provide meeting info via webex object or app context
    try {
      // Some embedded app contexts expose webex.meetings or similar
      if (app.about && app.about.meeting) {
        detectedMeetingInfo = app.about.meeting;
        log("Auto-detected meeting", detectedMeetingInfo);
        els.meetingDestination.value = detectedMeetingInfo.id || detectedMeetingInfo.url || '';
        const hint = document.getElementById("meetingDestinationHint");
        if (hint) {
          hint.textContent = "✓ Meeting auto-detected from embedded context.";
        }
      }
    } catch (e) {
      // Meeting detection failed; meeting destination is optional and can be manually provided
      log("Could not auto-detect meeting", e && e.message);
    }
  }

  // ---------- Webex JS SDK ----------

  function describeCredentials(sdk) {
    try {
      const creds = sdk && sdk.credentials;
      const st = creds && creds.supertoken;
      return {
        hasCredentials: Boolean(creds),
        canAuthorize: Boolean(sdk && sdk.canAuthorize),
        credentialsCanAuthorize: Boolean(creds && creds.canAuthorize),
        hasSupertoken: Boolean(st),
        supertokenHasAccessToken: Boolean(st && st.access_token),
        supertokenCanAuthorize: Boolean(st && st.canAuthorize),
        supertokenIsExpired: Boolean(st && st.isExpired),
      };
    } catch (e) {
      return { error: e && e.message };
    }
  }

  // Forces the supertoken onto the credentials plugin. The constructor normally
  // does this from `credentials.access_token`, but if storage rehydration or a
  // prior failed init leaves canAuthorize false, set it explicitly.
  function forceSupertoken(sdk, token) {
    try {
      if (sdk && sdk.credentials && typeof sdk.credentials.set === "function") {
        log("Forcing supertoken onto credentials plugin");
        sdk.credentials.set({
          supertoken: { access_token: token, token_type: "Bearer" },
        });
        return true;
      }
    } catch (e) {
      log("forceSupertoken failed", e && e.message);
    }
    return false;
  }

  function waitForSdkReady(sdk, token, timeoutMs = 15000) {
    return new Promise((resolve, reject) => {
      const tryAuthorize = (source) => {
        if (sdk.canAuthorize) {
          log("SDK ready", `canAuthorize=true (${source})`);
          return true;
        }
        return false;
      };

      // If already authorizable, resolve immediately.
      if (tryAuthorize("immediate")) {
        resolve();
        return;
      }

      let settled = false;
      const finish = (ok, reason) => {
        if (settled) return;
        settled = true;
        clearInterval(poll);
        clearTimeout(timer);
        if (ok) resolve();
        else reject(new Error(reason || "SDK cannot authorize"));
      };

      const resolveOrForce = (source) => {
        if (tryAuthorize(source)) {
          finish(true);
          return;
        }
        // Not authorizable yet: log state and try forcing the supertoken.
        log("Credentials state", describeCredentials(sdk));
        if (forceSupertoken(sdk, token) && tryAuthorize("forced")) {
          finish(true);
        }
      };

      // Preferred: wait for the 'ready' event.
      if (typeof sdk.once === "function") {
        sdk.once("ready", () => {
          log("SDK event", "ready");
          resolveOrForce("ready");
        });
      }

      // Fallback: poll canAuthorize in case the event already fired.
      const poll = setInterval(() => resolveOrForce("poll"), 250);

      const timer = setTimeout(() => {
        if (!sdk.canAuthorize) {
          log("Credentials state (timeout)", describeCredentials(sdk));
          forceSupertoken(sdk, token);
        }
        finish(
          sdk.canAuthorize,
          "Timed out waiting for SDK to authorize (token may be invalid for Meetings)"
        );
      }, timeoutMs);
    });
  }

  async function registerSdk() {
    if (typeof window.Webex === "undefined" || !window.Webex.init) {
      alert(
        "Webex JS SDK failed to load. Check your network or the CDN URL in index.html."
      );
      return;
    }

    setStatus("Registering...", "pending");
    try {
      // Preferred path: the OAuth-authorized SDK from the Webex login.
      if (webexSdk && webexSdk.canAuthorize) {
        log("Register step", "using signed-in OAuth session");
      } else {
        // Fallback: a pasted personal access token (testing only).
        const token = normalizeToken(els.accessToken.value);
        if (!token) {
          alert("Sign in with Webex first (or paste a personal token under Advanced).");
          setStatus("Idle", "idle");
          return;
        }
        const me = await validateAccessToken(token);
        log("Token valid for user", me && me.emails ? me.emails[0] : me && me.id);
        webexSdk = window.Webex.init({
          credentials: { access_token: token },
        });
        log("Register step", "waiting for SDK ready");
        await waitForSdkReady(webexSdk, token);
      }

      try {
        log("Register step", "meetings.register");
        await webexSdk.meetings.register();
      } catch (registerErr) {
        log("meetings.register failed", registerErr && (registerErr.message || registerErr));

        // Fallback for environments where meetings.register fails before device registration.
        if (webexSdk.internal && webexSdk.internal.device && webexSdk.internal.device.register) {
          log("Fallback step", "internal.device.register");
          await webexSdk.internal.device.register();
        } else {
          throw registerErr;
        }
      }

      log("Register step", "meetings.syncMeetings");
      await webexSdk.meetings.syncMeetings();

      log("Webex SDK registered & meetings synced");
      setStatus("Registered", "ok");
      els.btnJoin.disabled = false;
    } catch (err) {
      log("registerSdk failed", err && (err.message || err));
      log("registerSdk details", safeJson(err));
      if (String(err && err.message).includes("Token check failed (401)")) {
        log("Hint", "Use a fresh Personal Access Token from developer.webex.com (12h expiry).");
      } else if (String(err && err.message).includes("SDK cannot authorize")) {
        log("Hint", "Token is valid for REST but not accepted by Meetings auth. Try a token from the same Webex org/account currently in the meeting client.");
      }
      setStatus("Register failed", "err");
    }
  }

  async function joinMeeting() {
    if (!webexSdk) {
      alert("Register the SDK first.");
      return;
    }
    let destination = els.meetingDestination.value.trim();
    if (!destination) {
      alert("Meeting destination not detected or provided. Paste the meeting SIP URI, number, or link.");
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
    // Forwarding disabled in this development branch. Left intentionally
    // empty so the app only displays live transcription in the UI.
  }

  // ---------- wire up ----------

  function init() {
    restoreInputs();
    // Show the redirect URI the user must register on their Integration.
    els.redirectUri.value = getRedirectUri();
    // backend forwarding removed; only persist meetingDestination + clientId
    els.meetingDestination.addEventListener("change", persistInputs);
    els.clientId.addEventListener("change", persistInputs);

    // Register requires a signed-in session first.
    els.btnRegister.disabled = true;

    // Allow the Advanced personal-token fallback to enable Register too.
    els.accessToken.addEventListener("input", () => {
      if (els.accessToken.value.trim()) els.btnRegister.disabled = false;
    });

    els.btnLogin.addEventListener("click", loginWithWebex);
    els.btnLogout.addEventListener("click", logout);
    els.btnRegister.addEventListener("click", registerSdk);
    els.btnJoin.addEventListener("click", joinMeeting);
    els.btnLeave.addEventListener("click", leaveMeeting);

    initEmbeddedFramework();
    log("App initialized", "build v9");

    // If returning from a Webex login redirect, pick up the token.
    restoreOAuthSession();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
