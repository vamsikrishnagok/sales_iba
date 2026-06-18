"""Sample FastAPI backend for the Webex transcript app.

Run locally:
    pip install -r requirements.txt
    uvicorn main:app --reload --port 8000

Then the frontend (app.js) can POST transcript lines to /transcripts
and fetch a health check from /health.
"""

from datetime import datetime, timezone

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI(title="Webex Transcript API")

# Allow the browser app (GitHub Pages + localhost) to call this API.
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://vamsikrishnagok.github.io",
        "http://localhost:8000",
        "http://127.0.0.1:8000",
        "http://localhost:5500",
        "http://127.0.0.1:5500",
    ],
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory store (resets on restart). Fine for a demo.
transcripts: list[dict] = []


class TranscriptIn(BaseModel):
    """Shape of the JSON the frontend sends."""

    speaker: str
    text: str
    is_final: bool = False
    transcript_id: str | None = None


@app.get("/health")
def health() -> dict:
    """Simple endpoint app.js can call to confirm the backend is up."""
    return {"status": "ok", "count": len(transcripts)}


@app.post("/transcripts")
def add_transcript(item: TranscriptIn) -> dict:
    """Receive a transcript line from the frontend and store it."""
    record = {
        "speaker": item.speaker,
        "text": item.text,
        "is_final": item.is_final,
        "transcript_id": item.transcript_id,
        "received_at": datetime.now(timezone.utc).isoformat(),
    }
    transcripts.append(record)
    return {"saved": True, "total": len(transcripts)}


@app.get("/transcripts")
def list_transcripts() -> list[dict]:
    """Return everything received so far."""
    return transcripts
