"""FastAPI backend that receives Webex live transcription chunks and prints them.

Run locally:
    pip install -r requirements.txt
    uvicorn main:app --reload --host 0.0.0.0 --port 8000

POST a transcript chunk:
    curl -X POST http://localhost:8000/transcripts \
         -H "Content-Type: application/json" \
         -d '{"id":"x","personId":"y","transcription":"hello","timestamp":"2026-01-01T00:00:00Z","type":"final"}'
"""

from datetime import datetime
from typing import Optional

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI(title="Webex Live Transcript Receiver", version="0.1.0")

# CORS: the embedded app runs from GitHub Pages (a different origin), so the
# browser must be allowed to POST here. Tighten `allow_origins` to your
# GitHub Pages URL in production.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)


class TranscriptChunk(BaseModel):
    id: Optional[str] = None
    personId: Optional[str] = None
    transcription: str
    timestamp: Optional[str] = None
    type: Optional[str] = "final"  # "final" or "interim"
    meetingId: Optional[str] = None


@app.get("/")
def root():
    return {"service": "webex-transcript-receiver", "status": "ok"}


@app.get("/health")
def health():
    return {"status": "ok", "time": datetime.utcnow().isoformat() + "Z"}


@app.post("/transcripts")
def receive_transcript(chunk: TranscriptChunk):
    received_at = datetime.utcnow().isoformat() + "Z"
    print(
        f"[{received_at}] "
        f"type={chunk.type} "
        f"meetingId={chunk.meetingId} "
        f"personId={chunk.personId} "
        f"ts={chunk.timestamp} :: "
        f"{chunk.transcription}",
        flush=True,
    )
    return {"ok": True, "received_at": received_at}
