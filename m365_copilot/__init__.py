"""M365 Copilot Python Client."""

from .auth import get_token, get_token_info, decode_jwt, token_from_browser_js
from .signalr import chat, ChatResponse, build_ws_url
from .session import CopilotSession, MODEL_TONES

__all__ = [
    "get_token", "get_token_info", "decode_jwt",
    "chat", "ChatResponse", "build_ws_url",
    "CopilotSession", "MODEL_TONES",
]
