# Webex Live Transcript &rarr; FastAPI

A Webex Embedded App that captures **real-time meeting transcription** via the
Webex JavaScript SDK and displays transcripts live inside the embedded app.
Forwarding to the backend is disabled in this branch — transcripts are shown
in the UI only.

---

## Architecture

```
+--------------------+        HTTPS (POST JSON)        +---------------------+
|  Webex Embedded    |  ----------------------------> |  FastAPI Backend    |
|  App (this repo,   |                                |  (prints to stdout) |
|  hosted on GH      |                                |                     |
|  Pages)            |                                +---------------------+
|                    |
|  Loads:            |
|  - Embedded App    |
|    Framework SDK   |
|  - Webex JS SDK    |
|                    |
|  Subscribes to:    |
|  meeting:receive   |
|  Transcription:    |
|  started           |
+--------------------+
```

The page uses two distinct Webex SDKs:

1. **Webex Embedded App Framework SDK** &mdash; required so the page can run
   inside the Webex client. Loaded from
   `https://binaries.webex.com/static-content-pipeline/webex-embedded-app/v1/webex-embedded-app-sdk.js`.
2. **Webex JavaScript SDK** &mdash; used to register a Webex user device,
   join the meeting with `receiveTranscription: true`, and subscribe to
   `meeting:receiveTranscription:started` events (per the
   [official blog](https://developer.webex.com/blog/how-to-receive-real-time-meeting-transcription-with-the-webex-javascript-sdk)).

> **Note** &mdash; the Embedded App Framework itself does **not** expose a
> transcription API; that is why we also load the Webex JS SDK and have it
> join the meeting alongside the user.

---

## Repository layout

```
.
├── index.html              # GitHub Pages entry point
├── css/style.css
├── js/app.js               # SDK init + transcription display (no forwarding)
├── backend/
│   ├── main.py             # FastAPI app
│   └── requirements.txt
├── .gitignore
└── README.md
```

---

## Prerequisites

* A Webex account that is **not** a free consumer account &mdash; embedded apps
  require a paid/developer-sandbox org. Request one via
  [Webex Developer Sandbox](https://developer.webex.com/docs/developer-sandbox-guide).
* The host of the meeting must have **Webex Assistant** enabled so that
  transcription events are produced.
* A Webex personal access token from
  [developer.webex.com/docs/getting-started](https://developer.webex.com/docs/getting-started)
  (valid for 12 hours). For long-running use, build an OAuth integration.
* Python 3.10+ and `pip` for the backend.
* A GitHub account (this repo is already at
  `https://github.com/vamsikrishnagok/sales_iba`).

---

## Part 1 &mdash; Run the FastAPI backend

The backend just receives JSON payloads and prints them to stdout.

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

Test it with a manual request:

```powershell
curl -X POST http://localhost:8000/transcripts `
     -H "Content-Type: application/json" `
     -d '{"transcription":"hello world","type":"final"}'
```

You should see something like this in the uvicorn console:

```
[2026-06-17T12:34:56.789Z] type=final meetingId=None personId=None ts=None :: hello world
```

### Exposing the backend to the browser

The browser running the embedded app (inside the Webex client) must be able to
reach the backend over the network. The two simplest options are:

* **Same machine** &mdash; use `http://localhost:8000/transcripts` (works only
  when you are testing the embedded app locally on the same machine).
* **Public HTTPS tunnel** &mdash; recommended. Use
  [ngrok](https://ngrok.com/) or
  [Cloudflare Tunnel](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/):

  ```powershell
  ngrok http 8000
  ```

  Use the resulting `https://xxxx.ngrok.io/transcripts` URL inside the app.

> The GitHub Pages site is served over HTTPS, and browsers block mixed
> content. **The backend URL you paste into the app must therefore be
> HTTPS**, not plain `http://`, unless you are running everything on
> `localhost`.

---

## Part 2 &mdash; Publish the frontend on GitHub Pages

1. Push this repository to GitHub (already done for
   `https://github.com/vamsikrishnagok/sales_iba`).
2. In GitHub, go to **Settings &rarr; Pages**.
3. Under **Build and deployment**:
   * Source: **Deploy from a branch**
   * Branch: `main`, folder: `/ (root)`
4. Save. After ~1 minute the site is published at:

   ```
   https://vamsikrishnagok.github.io/sales_iba/
   ```

5. Verify the URL loads in a normal browser. You should see the
   "Live Transcript &rarr; FastAPI" UI. (Outside Webex, the "Embedded App
   Context" card will say the framework is not available &mdash; that is
   expected.)

---

## Part 3 &mdash; Register the Embedded App on Webex

1. Go to
   [developer.webex.com &rarr; My Apps &rarr; Create a New App](https://developer.webex.com/my-apps/new/embedded-app).
2. Choose **Embedded App**.
3. Fill in:
   * **App name**: e.g. `Live Transcript Forwarder`
   * **Description**, **icons**, etc.
   * **Where can your app be used?** check at least **Meetings** (and
     optionally **Spaces**).
   * **Valid domains**: add `vamsikrishnagok.github.io` (the host of the
     start page).
   * **Start Page URL**: `https://vamsikrishnagok.github.io/sales_iba/`
   * **Privacy policy URL / Support URL**: any valid URL you control.
4. Set the visibility (initially **Private** is easiest for development).
5. Save and submit.

> Webex requires a public HTTPS URL for the Start Page. GitHub Pages
> satisfies this requirement.

---

## Part 4 &mdash; Use the app during a meeting

1. Join (or start) a Webex meeting in the **desktop** Webex client.
2. Make sure **Webex Assistant is ON** (host enables it).
3. Click the **Apps** button in the meeting toolbar and open
   "Live Transcript Forwarder".
4. In the app UI:
  1. Paste your **Webex personal access token**.
  2. Paste the **meeting destination** &mdash; this is typically the
    meeting SIP URI (e.g. `123456789@webex.com`) or the personal-room
    link of the host.
  3. Click **Register Webex SDK**, then **Join &amp; Start Transcription**.
5. Watch live transcript lines appear in the **Live Transcript** card and in
   the uvicorn console.

> The Webex JS SDK joins as a *separate* participant on behalf of the user
> whose token you supplied. That participant must be admitted to the meeting
> like any other attendee.

---

## Payload format

The browser POSTs JSON of the following shape to your backend:

```json
{
  "id": "abc-123",
  "personId": "Y2lzY29zcGFyazov...",
  "transcription": "Hello everyone, thanks for joining.",
  "timestamp": "2026-06-17T12:34:56.789Z",
  "type": "final",
  "meetingId": "0bcd..."
}
```

`type` is either `"final"` or `"interim"` &mdash; mirroring the upstream
Webex transcription event.

---

## Security notes

* **Never commit your access token.** The app keeps it in the input field
  (in-memory) and `localStorage` is *not* used for the token.
* **Restrict CORS** on the backend (`backend/main.py`) to your GitHub Pages
  origin (`https://vamsikrishnagok.github.io`) before exposing it publicly.
* Webex transcripts may contain sensitive content &mdash; secure the backend
  with authentication (e.g. an API key header) before deploying anywhere
  other than localhost / dev.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| "Webex Embedded App SDK not detected" | Page is loaded outside Webex. Open the app inside a Webex meeting. |
| `registerSdk` fails with 401 | Access token expired (tokens from the docs portal last 12 hours). Get a fresh one. |
| No transcript events arrive | Webex Assistant not turned on for the meeting, or the SDK participant is in the lobby. |
| Browser console shows CORS errors | Set `allow_origins` in `backend/main.py` to include your Pages URL, or use a tunnel that returns proper CORS headers. |
| "Mixed Content" warning | Backend URL must be HTTPS when the page is served from GitHub Pages. |

---

## References

* [Webex Embedded Apps overview](https://developer.webex.com/create/docs/embedded-apps)
* [How to Receive Real-Time Meeting Transcription with the Webex JS SDK](https://developer.webex.com/blog/how-to-receive-real-time-meeting-transcription-with-the-webex-javascript-sdk)
* [Webex JS SDK on GitHub](https://github.com/webex/webex-js-sdk)
* [FastAPI docs](https://fastapi.tiangolo.com/)
