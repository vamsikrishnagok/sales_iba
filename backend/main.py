"""Sample FastAPI backend for the Webex transcript app.

Run locally:
    pip install -r requirements.txt
    uvicorn main:app --reload --port 8000

Then the frontend (app.js) can POST transcript lines to /transcripts
and fetch a health check from /health.
"""

import json
from datetime import datetime, timezone
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from llm.provider import ask_llm

app = FastAPI(title="Webex Transcript API")

# Directory where completed-meeting transcripts are written.
TRANSCRIPTS_DIR = Path(__file__).parent / "transcripts"
TRANSCRIPTS_DIR.mkdir(exist_ok=True)

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

# Live session file, continuously updated as lines arrive so a sudden
# meeting close (which may skip the /meeting/end call) never loses data.
SESSION_FILE = TRANSCRIPTS_DIR / "current-session.json"


def _write_session_file() -> None:
    """Persist the current in-memory transcript to the live session file."""
    SESSION_FILE.write_text(
        json.dumps(transcripts, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def _finalize_session() -> dict:
    """Write timestamped .txt/.json files and clear the live session."""
    if not transcripts:
        return {"saved": False, "reason": "no transcripts", "lines": 0}

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    base = TRANSCRIPTS_DIR / f"transcript-{timestamp}"

    # Human-readable text file.
    lines = [f"{item['speaker']}: {item['text']}" for item in transcripts]
    text_path = base.with_suffix(".txt")
    text_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    # Structured JSON alongside it for later processing.
    json_path = base.with_suffix(".json")
    json_path.write_text(
        json.dumps(transcripts, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    saved_lines = len(transcripts)
    transcripts.clear()
    # Reset the live session file now that it's been finalized.
    if SESSION_FILE.exists():
        SESSION_FILE.unlink()
    return {
        "saved": True,
        "lines": saved_lines,
        "text_file": str(text_path),
        "json_file": str(json_path),
    }


class TranscriptIn(BaseModel):
    """Shape of the JSON the frontend sends."""

    speaker: str
    text: str
    is_final: bool = False
    transcript_id: str | None = None


class AskIn(BaseModel):
    """A question the frontend wants the LLM to answer."""

    question: str
    speaker: str | None = None


class TranscriptLine(BaseModel):
    """One finalized transcript line sent by the frontend on save."""

    speaker: str
    text: str
    transcript_id: str | None = None


class SaveTranscriptIn(BaseModel):
    """Full transcript the frontend sends when the user clicks Save."""

    lines: list[TranscriptLine]


@app.on_event("startup")
def recover_orphaned_session() -> None:
    """If a previous run ended abruptly, finalize its leftover session file."""
    if SESSION_FILE.exists():
        try:
            data = json.loads(SESSION_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            data = []
        if data:
            transcripts.extend(data)
            result = _finalize_session()
            print(f"[recovery] finalized orphaned session: {result}")
        else:
            SESSION_FILE.unlink(missing_ok=True)


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
    # Persist immediately so nothing is lost if the meeting closes abruptly.
    _write_session_file()
    return {"saved": True, "total": len(transcripts)}


@app.get("/transcripts")
def list_transcripts() -> list[dict]:
    """Return everything received so far."""
    return transcripts


@app.post("/meeting/end")
def end_meeting() -> dict:
    """Persist the full transcript to a file and reset the in-memory store."""
    return _finalize_session()


@app.post("/transcripts/save")
def save_transcript(payload: SaveTranscriptIn) -> dict:
    """Save a complete transcript sent from the frontend to timestamped files."""
    if not payload.lines:
        raise HTTPException(status_code=400, detail="No transcript lines to save.")

    now = datetime.now(timezone.utc)
    records = [
        {
            "speaker": line.speaker,
            "text": line.text,
            "transcript_id": line.transcript_id,
        }
        for line in payload.lines
    ]

    timestamp = now.strftime("%Y%m%d-%H%M%S")
    base = TRANSCRIPTS_DIR / f"transcript-{timestamp}"

    text_lines = [f"{r['speaker']}: {r['text']}" for r in records]
    text_path = base.with_suffix(".txt")
    text_path.write_text("\n".join(text_lines) + "\n", encoding="utf-8")

    json_path = base.with_suffix(".json")
    json_path.write_text(
        json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    return {
        "saved": True,
        "lines": len(records),
        "text_file": str(text_path),
        "json_file": str(json_path),
    }


@app.post("/ask")
def ask(item: AskIn) -> dict:
    """Answer a question (a transcript line containing '?') using the LLM."""
    question = item.question.strip()
    if not question:
        raise HTTPException(status_code=400, detail="Question is empty.")
    try:
        answer = ask_llm(question)
    except Exception as exc:  # surface config/credential errors to the client
        raise HTTPException(status_code=502, detail=f"LLM error: {exc}") from exc
    return {
        "question": question,
        "speaker": item.speaker,
        "answer": answer,
        "answered_at": datetime.now(timezone.utc).isoformat(),
    }
