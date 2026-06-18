"""LLM provider: Gemini 2.5 Pro via Vertex AI service-account credentials.

Mirrors the manager-provided pattern (service-account JSON fields supplied
through env vars). Returns a LangChain ``BaseChatModel`` so it slots into
``create_agent(model=...)`` unchanged.
"""
from __future__ import annotations

from functools import lru_cache

from google.oauth2 import service_account
from langchain_core.language_models import BaseChatModel
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_google_vertexai import ChatVertexAI

from config import settings


@lru_cache(maxsize=1)
def _get_credentials() -> service_account.Credentials:
    if not (settings.gcp_project_id and settings.gcp_client_email and settings.gcp_private_key):
        raise RuntimeError(
            "Missing Vertex AI credentials. Set GCP_PROJECT_ID, GCP_CLIENT_EMAIL, "
            "and GCP_PRIVATE_KEY in backend/.env."
        )
    return service_account.Credentials.from_service_account_info(
        {
            "project_id": settings.gcp_project_id,
            "client_email": settings.gcp_client_email,
            "private_key": settings.gcp_private_key,
            "token_uri": "https://oauth2.googleapis.com/token",
        },
        scopes=["https://www.googleapis.com/auth/cloud-platform"],
    )


@lru_cache(maxsize=1)
def get_chat_model() -> BaseChatModel:
    provider = settings.llm_provider.lower()
    if provider == "vertex":
        # Defer credential validation to first call so imports stay cheap.
        return ChatVertexAI(
            model=settings.gemini_model,
            project=settings.gcp_project_id or None,
            location=settings.gcp_location,
            credentials=_get_credentials() if settings.gcp_private_key else None,
            temperature=0.1,
        )
    raise NotImplementedError(
        f"LLM provider '{provider}' not wired. Implement in app/llm/provider.py."
    )


_SYSTEM_PROMPT = (
    "You are a helpful assistant answering questions asked aloud during a live "
    "meeting. Answer clearly and concisely in a few sentences."
)


def ask_llm(question: str) -> str:
    """Send a question to the LLM and return its text answer."""
    model = get_chat_model()
    response = model.invoke(
        [
            SystemMessage(content=_SYSTEM_PROMPT),
            HumanMessage(content=question),
        ]
    )
    # ChatVertexAI returns an AIMessage; content may be str or list of parts.
    content = response.content
    if isinstance(content, list):
        return "".join(
            part if isinstance(part, str) else part.get("text", "")
            for part in content
        )
    return content
