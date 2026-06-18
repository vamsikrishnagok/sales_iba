"""Application configuration loaded from environment variables."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    database_path: Path = Path(os.getenv("DATABASE_PATH", "./data/analytics.db")).resolve()
    llm_provider: str = os.getenv("LLM_PROVIDER", "vertex")
    gemini_model: str = os.getenv("GEMINI_MODEL", "gemini-2.5-pro")
    gcp_project_id: str = os.getenv("GCP_PROJECT_ID", "")
    gcp_client_email: str = os.getenv("GCP_CLIENT_EMAIL", "")
    gcp_private_key: str = os.getenv("GCP_PRIVATE_KEY", "").replace("\\n", "\n")
    gcp_location: str = os.getenv("GCP_LOCATION", "us-central1")
    copilotkit_endpoint: str = os.getenv("COPILOTKIT_ENDPOINT", "/copilotkit")


settings = Settings()
