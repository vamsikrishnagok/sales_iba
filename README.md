# Webex Live Meeting Transcription

A browser app that logs into Webex via OAuth, joins your current meeting, and
displays the real-time transcription. Finalized lines are forwarded to an
optional FastAPI backend.

**Live demo:** https://vamsikrishnagok.github.io/sales_iba/

## How it works

1. You log in with Webex (OAuth Integration).
2. The app registers the Webex Meetings SDK and finds your active meeting.
3. It joins and subscribes to live transcription events.
4. Each utterance is rendered as a single line prefixed with the speaker's name.
5. Finalized lines are POSTed to the FastAPI backend (if running).

## Project structure

```
index.html        Frontend page
app.js            Frontend logic (Webex SDK + transcription + backend calls)
style.css         Styles
webex.min.js      Webex JavaScript SDK
meetings.min.js   Webex Meetings SDK
backend/
  main.py         FastAPI backend (sample endpoints)
  requirements.txt
```

## Frontend setup

The frontend is static — just open `index.html`, or serve it locally:

```powershell
# from the project root
python -m http.server 5500
# then visit http://localhost:5500
```

### Webex OAuth configuration

In **developer.webex.com → My Webex Apps → your Integration**:

- **Redirect URI:** `https://vamsikrishnagok.github.io/sales_iba/`
  (must match `REDIRECT_URI` in `app.js` exactly, including the trailing slash)
- **Scopes:** `spark:all`, `spark:kms`
- Copy the **Client ID** into `CLIENT_ID` in `app.js`.

## Backend setup

```powershell
cd backend
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

Interactive API docs: http://localhost:8000/docs

### Endpoints

| Method | Path           | Description                          |
| ------ | -------------- | ------------------------------------ |
| GET    | `/health`      | Health check used by the frontend    |
| POST   | `/transcripts` | Store one transcript line            |
| GET    | `/transcripts` | List all stored lines                |

The frontend's `API_BASE` in `app.js` points to `http://localhost:8000`.

> **Note:** GitHub Pages is served over HTTPS, so browsers block calls to an
> `http://localhost` backend (mixed content). Run the frontend locally over
> HTTP for backend calls, or deploy the backend over HTTPS for production.

## Notes

- The backend store is in-memory and resets on restart (demo only).
- Always open the site at the exact registered redirect URI to avoid
  `redirect_uri_mismatch` OAuth errors.
