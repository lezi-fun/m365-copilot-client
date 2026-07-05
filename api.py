"""
M365 Copilot OpenAI-Compatible API Server

Usage:
    export M365_TOKEN='***'
    python m365_api.py --port 23100

Or with a config token:
    python m365_api.py
    # Uses M365_TOKEN env, file cache, or MSAL auth
"""

import json
import logging
import os
import sys
import time
import uuid
from pathlib import Path
from typing import AsyncGenerator, Optional

import uvicorn
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

# Add parent dir so m365_copilot import works
sys.path.insert(0, str(Path(__file__).parent))

from m365_copilot import (
    get_token,
    get_token_info,
    token_from_browser_js,
    CopilotSession,
)

logger = logging.getLogger("m365.api")

# --- App Setup ---

app = FastAPI(title="M365 Copilot API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Token Resolution ---

_api_token: Optional[str] = None
_TOKEN_FILE = Path.home() / ".config" / "m365-copilot" / "token.txt"

def resolve_token() -> str:
    """Resolve the M365 token from env var, file, or MSAL auth."""
    global _api_token
    if _api_token:
        return _api_token

    # 1. Env var
    token = os.environ.get("M365_TOKEN")
    if token:
        _api_token = token
        logger.info("Token from M365_TOKEN env var")
        return token

    # 2. Token file (from m365-copilot auth or manual save)
    if _TOKEN_FILE.exists():
        try:
            token = _TOKEN_FILE.read_text().strip()
            if token:
                _api_token = token
                logger.info("Token from %s", _TOKEN_FILE)
                return token
        except Exception:
            pass

    # 3. Browser JS
    token = token_from_browser_js()
    if token:
        _api_token = token
        return token

    # 4. MSAL cache
    token = get_token()
    if token:
        _api_token = token
        logger.info("Token from MSAL cache")
        return token

    raise HTTPException(
        status_code=401,
        detail="No M365 token. Set M365_TOKEN or run auth first.",
    )

# --- Model Map ---

AVAILABLE_MODELS = {
    "auto": "magic",
    "quick": "Gpt_Quick",
    "think-deeper": "Gpt_Reasoning",
    "claude-sonnet": "Claude_Sonnet",
    "claude-sonnet-4.6": "Claude_Sonnet",
    "claude-opus": "Claude_Opus",
    "gpt-5.5": "Gpt_5_5_Chat",
    "gpt-5.5-quick": "Gpt_5_5_Chat",
    "gpt-5.5-think-deeper": "Gpt_5_5_Reasoning",
    "gpt-5.4": "Gpt_5_4_Reasoning",
    "gpt-5.4-quick": "Gpt_5_4_Quick",
    "gpt-5.3": "Gpt_5_3_Quick",
    "gpt-5.2": "Gpt_5_2_Quick",
}

def get_tone(model: str) -> str:
    return AVAILABLE_MODELS.get(model) or "magic"

# --- Request Models ---

class ChatMessage(BaseModel):
    role: str
    content: str

class ChatCompletionRequest(BaseModel):
    model: str = "auto"
    messages: list[ChatMessage]
    stream: bool = False
    user: Optional[str] = None

class ModelInfo(BaseModel):
    id: str
    object: str = "model"
    created: int = 0
    owned_by: str = "m365-copilot"

class ModelList(BaseModel):
    object: str = "list"
    data: list[ModelInfo]

# --- Helpers ---

def build_chat_id() -> str:
    return f"chatcmpl-{uuid.uuid4().hex[:16]}"

def build_openai_chunk(chat_id: str, model: str, content: str,
                       finish_reason: Optional[str] = None) -> str:
    chunk = {
        "id": chat_id,
        "object": "chat.completion.chunk",
        "created": int(time.time()),
        "model": model,
        "choices": [{
            "index": 0,
            "delta": {"content": content} if content else {},
            "finish_reason": finish_reason,
        }],
    }
    return f"data: {json.dumps(chunk, ensure_ascii=False)}\n\n"

def build_openai_response(chat_id: str, model: str, content: str,
                          usage: Optional[dict] = None) -> dict:
    resp = {
        "id": chat_id,
        "object": "chat.completion",
        "created": int(time.time()),
        "model": model,
        "choices": [{
            "index": 0,
            "message": {"role": "assistant", "content": content},
            "finish_reason": "stop",
        }],
    }
    if usage:
        resp["usage"] = usage
    return resp

# --- Session Pool with message-based context matching ---

from collections import OrderedDict
import hashlib

_session_pool: OrderedDict[str, CopilotSession] = OrderedDict()
_msg_prefix_to_session: dict[str, str] = {}
_MAX_SESSIONS = 100

def _messages_fingerprint(messages: list[ChatMessage], count: int) -> str:
    """Hash the first `count` messages to create a fingerprint."""
    prefix = messages[:count]
    data = json.dumps(
        [{"role": m.role, "content": m.content[:300]} for m in prefix],
        sort_keys=True, ensure_ascii=False,
    )
    return hashlib.sha256(data.encode()).hexdigest()[:16]

def get_or_create_session(token: str, messages: list[ChatMessage]) -> tuple[CopilotSession, str]:
    """
    Match conversation by message history prefix.
    - If messages[0:-1] matches a previous session → reuse it
    - Otherwise → create new session
    
    Returns (session, session_key).
    """
    # Try to find an existing session by message prefix (all but last)
    if len(messages) > 1:
        prefix_key = _messages_fingerprint(messages, len(messages) - 1)
        if prefix_key in _msg_prefix_to_session:
            session_key = _msg_prefix_to_session[prefix_key]
            if session_key in _session_pool:
                # Move to end (most recently used)
                session = _session_pool.pop(session_key)
                _session_pool[session_key] = session
                logger.debug("Reused session: key=%s turn=%d", session_key[:8], session.turn_count)
                return session, session_key

    # New session
    session_key = str(uuid.uuid4())
    session = CopilotSession(token=token)
    _session_pool[session_key] = session

    # LRU eviction
    while len(_session_pool) > _MAX_SESSIONS:
        old_key, old_session = _session_pool.popitem(last=False)
        # Also clean up prefix mappings pointing to this session
        keys_to_delete = [k for k, v in _msg_prefix_to_session.items() if v == old_key]
        for k in keys_to_delete:
            del _msg_prefix_to_session[k]

    # Store mapping from full message hash → session key
    full_key = _messages_fingerprint(messages, len(messages))
    _msg_prefix_to_session[full_key] = session_key

    logger.info("New session: key=%s conv=%s", session_key[:8], session.conversation_id[:8])
    return session, session_key


# --- Routes ---

@app.get("/v1/models")
async def list_models():
    return ModelList(data=[
        ModelInfo(id=name, created=int(time.time()))
        for name in AVAILABLE_MODELS
    ])


@app.post("/v1/chat/completions")
async def chat_completions(body: ChatCompletionRequest, request: Request):
    try:
        token = resolve_token()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=401, detail=str(e))

    if not body.messages:
        raise HTTPException(status_code=400, detail="messages is required")

    # Extract user messages
    last_user_msg = None
    for msg in reversed(body.messages):
        if msg.role == "user":
            last_user_msg = msg.content
            break
    if not last_user_msg:
        raise HTTPException(status_code=400, detail="No user message")

    # Conversation context: auto-match by message history
    session, active_conv_id = get_or_create_session(token, body.messages)

    tone = get_tone(body.model)

    logger.info("Chat: model=%s tone=%s turn=%d conv=%s msg=%s",
        body.model, tone, session.turn_count, active_conv_id[:8], last_user_msg[:80])

    chat_id = build_chat_id()

    async def stream_response() -> AsyncGenerator[str, None]:
        nonlocal chat_id
        full_text = ""

        def on_chunk(chunk: str):
            nonlocal full_text
            full_text += chunk

        resp = await session.send(last_user_msg, tone=tone, on_delta=on_chunk)

        if resp.disengaged:
            yield build_openai_chunk(chat_id, body.model, "")
            yield "data: [DONE]\n\n"
            return

        if resp.error:
            yield build_openai_chunk(chat_id, body.model, "")
            yield "data: [DONE]\n\n"
            return

        text = full_text or resp.text
        if text:
            yield build_openai_chunk(chat_id, body.model, text)

        yield build_openai_chunk(chat_id, body.model, "", finish_reason="stop")
        yield "data: [DONE]\n\n"

    resp_headers = {
        "X-Conversation-Id": active_conv_id,
    }

    if body.stream:
        return StreamingResponse(
            stream_response(),
            media_type="text/event-stream",
            headers={
                **resp_headers,
                "Cache-Control": "no-cache",
                "X-Accel-Buffering": "no",
            },
        )

    # Non-streaming
    full_text = ""
    def on_chunk(chunk: str):
        nonlocal full_text
        full_text += chunk

    resp = await session.send(last_user_msg, tone=tone, on_delta=on_chunk)

    if resp.error:
        raise HTTPException(status_code=502, detail=resp.error)
    if resp.disengaged:
        raise HTTPException(status_code=502, detail="Disengaged by M365")

    text = full_text or resp.text
    return build_openai_response(chat_id, body.model, text, usage={
        "prompt_tokens": len(last_user_msg),
        "completion_tokens": len(text),
        "total_tokens": len(last_user_msg) + len(text),
    })


@app.get("/health")
async def health():
    return {"status": "ok", "models": list(AVAILABLE_MODELS.keys())}


# --- Main ---

def main():
    import argparse

    parser = argparse.ArgumentParser(description="M365 Copilot API Server")
    parser.add_argument("--host", default="127.0.0.1", help="Bind address")
    parser.add_argument("--port", type=int, default=23100, help="Bind port")
    parser.add_argument("--reload", action="store_true", help="Auto-reload")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose")
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger("m365").setLevel(logging.DEBUG)

    try:
        token = resolve_token()
        info = get_token_info(token)
        logger.info("Token: user=%s oid=%s", info.get("name", "?"), info.get("oid", "?")[:8])
    except Exception as e:
        logger.warning("No token: %s", e)

    uvicorn.run(
        "api:app",
        host=args.host,
        port=args.port,
        reload=args.reload,
    )


if __name__ == "__main__":
    main()
