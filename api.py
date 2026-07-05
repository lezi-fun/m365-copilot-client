"""
M365 Copilot OpenAI-Compatible API Server — with tool calling support.
"""

import json
import logging
import os
import re
import sys
import time
import uuid
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
from typing import Any, AsyncGenerator, Optional

import hashlib
import uvicorn
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel

from m365_copilot import (
    get_token,
    get_token_info,
    token_from_browser_js,
    CopilotSession,
)
import uvicorn
from fastapi import FastAPI, HTTPException, Request

logger = logging.getLogger("m365.api")

app = FastAPI(title="M365 Copilot API", version="0.2.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# --- Token ---

_api_token: Optional[str] = None
_TOKEN_FILE = Path.home() / ".config" / "m365-copilot" / "token.txt"

def resolve_token() -> str:
    global _api_token
    if _api_token: return _api_token

    # M365_TOKEN env var
    t = os.environ.get("M365_TOKEN")
    if t:
        _api_token = t; logger.info("Token: env"); return t

    # Token file
    if _TOKEN_FILE.exists():
        t = _TOKEN_FILE.read_text().strip()
        if t:
            _api_token = t; logger.info("Token: file"); return t

    # Browser JS / MSAL (lazy, only if needed)
    t = token_from_browser_js()
    if t:
        _api_token = t; logger.info("Token: browser"); return t

    t = get_token()
    if t:
        _api_token = t; logger.info("Token: MSAL"); return t

    raise HTTPException(status_code=401, detail="No token")

def _token_status():
    """Return current token info or None."""
    global _api_token
    try:
        t = resolve_token()
        info = get_token_info(t)
        return {"user": info.get("name"), "expires_in": info.get("expires_in"), "ok": info.get("expires_in", 0) > 60}
    except:
        return None

# --- Models ---

MODEL_TONES = {
    "auto": "magic", "quick": "Gpt_Quick", "think-deeper": "Gpt_Reasoning",
    "claude-sonnet": "Claude_Sonnet", "claude-sonnet-4.6": "Claude_Sonnet",
    "claude-opus": "Claude_Opus",
    "gpt-5.5": "Gpt_5_5_Chat", "gpt-5.5-quick": "Gpt_5_5_Chat",
    "gpt-5.5-think-deeper": "Gpt_5_5_Reasoning",
    "gpt-5.4": "Gpt_5_4_Reasoning", "gpt-5.4-quick": "Gpt_5_4_Quick",
    "gpt-5.3": "Gpt_5_3_Quick", "gpt-5.2": "Gpt_5_2_Quick",
}

# --- Schemas ---

class ToolFunction(BaseModel):
    name: str; description: str = ""; parameters: dict[str, Any] = {}
class ToolDef(BaseModel):
    type: str = "function"; function: ToolFunction
class ChatMessage(BaseModel):
    role: str; content: str = ""
    tool_calls: Optional[list[dict]] = None
    tool_call_id: Optional[str] = None
class ChatCompletionRequest(BaseModel):
    model: str = "auto"; messages: list[ChatMessage]; stream: bool = False
    tools: Optional[list[ToolDef]] = None; tool_choice: Any = None
class ModelInfo(BaseModel):
    id: str; object: str = "model"; created: int = 0; owned_by: str = "m365-copilot"
class ModelList(BaseModel):
    object: str = "list"; data: list[ModelInfo]

# --- Built-in Tools ---

BUILTIN_TOOLS = [
    {"type": "function", "function": {
        "name": "get_current_time",
        "description": "获取当前日期和时间",
        "parameters": {"type": "object", "properties": {}, "required": []},
    }},
    {"type": "function", "function": {
        "name": "calculate",
        "description": "执行数学计算",
        "parameters": {"type": "object", "properties": {
            "expression": {"type": "string", "description": "数学表达式"},
        }, "required": ["expression"]},
    }},
]

def _execute_tool(name: str, args: dict) -> str:
    if name == "get_current_time":
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if name == "calculate":
        try:
            allowed = {"abs": abs, "round": round, "min": min, "max": max, "pow": pow}
            return str(eval(args["expression"], {"__builtins__": {}}, allowed))
        except Exception as e:
            return f"Error: {e}"
    return f"Unknown tool: {name}"

# Tool prompt injection
TOOL_PROMPT = """You have access to tools. When you need to use one, output EXACTLY this format on its own line:

```tool_call
{"tool": "tool_name", "arguments": {"arg1": "val1"}}
```

After getting the result, either call another tool or give the final answer in plain text."""

def _inject_tools(messages: list[dict], tools: list[dict]) -> list[dict]:
    if not tools:
        return messages
    prompt = TOOL_PROMPT + "\n\nAvailable tools:\n" + json.dumps(tools, indent=2, ensure_ascii=False)
    for m in messages:
        if m.get("role") == "system":
            m["content"] = prompt + "\n\n" + (m.get("content") or "")
            return messages
    return [{"role": "system", "content": prompt}] + messages

def _find_tool_calls(text: str) -> list[dict]:
    calls = []
    for m in re.finditer(r'```tool_call\s*\n(\{.*?\})\s*\n\s*```', text, re.DOTALL):
        try:
            calls.append(json.loads(m.group(1)))
        except json.JSONDecodeError:
            pass
    if calls:
        return calls
    # Bare JSON fallback
    for m in re.finditer(r'\{"tool"\s*:', text):
        try:
            start = m.start()
            depth = 0
            for i in range(start, len(text)):
                if text[i] == '{': depth += 1
                elif text[i] == '}':
                    depth -= 1
                    if depth == 0:
                        obj = json.loads(text[start:i+1])
                        if "tool" in obj and "arguments" in obj:
                            calls.append(obj)
                        break
        except (json.JSONDecodeError, ValueError):
            pass
    return calls

def _make_chat_id() -> str:
    return f"chatcmpl-{uuid.uuid4().hex[:16]}"

def _sse(data: dict) -> str:
    return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"

def _delta(model: str, cid: str, content: str = "", tcs: Optional[list] = None, fr: Optional[str] = None) -> str:
    d = {}
    if content: d["content"] = content
    if tcs: d["tool_calls"] = tcs
    return _sse({"id": cid, "object": "chat.completion.chunk", "created": int(time.time()),
                  "model": model, "choices": [{"index": 0, "delta": d, "finish_reason": fr}]})

# --- Session Pool ---

_session_pool: OrderedDict[str, CopilotSession] = OrderedDict()
_msg_map: dict[str, str] = {}
_MAX_S = 100

def _session(token: str, msgs: list[ChatMessage]) -> tuple[CopilotSession, str]:
    # 用第一条用户消息的内容作为会话 key
    first_user = ""
    for m in msgs:
        if m.role == "user":
            first_user = (m.content or "")[:200]
            break

    if first_user:
        k = hashlib.sha256(first_user.encode()).hexdigest()[:16]
        if k in _msg_map and _msg_map[k] in _session_pool:
            sk = _msg_map[k]
            s = _session_pool.pop(sk)
            _session_pool[sk] = s
            return s, sk

    sk = str(uuid.uuid4())
    s = CopilotSession(token=token)
    _session_pool[sk] = s
    while len(_session_pool) > _MAX_S:
        ok, _ = _session_pool.popitem(last=False)
        for k in list(_msg_map.keys()):
            if _msg_map[k] == ok:
                del _msg_map[k]
    if first_user:
        _msg_map[hashlib.sha256(first_user.encode()).hexdigest()[:16]] = sk
    return s, sk

# --- Routes ---

import os
_CHAT_HTML = os.path.join(os.path.dirname(__file__), "chat.html")

@app.get("/")
async def index():
    if os.path.exists(_CHAT_HTML):
        return FileResponse(_CHAT_HTML)
    return {"msg": "M365 Copilot API - POST /v1/chat/completions"}

@app.get("/v1/token/status")
async def token_status():
    info = _token_status()
    if info:
        return info
    raise HTTPException(status_code=401, detail="No valid token")

@app.post("/v1/token/refresh")
async def token_refresh():
    global _api_token
    _api_token = None
    info = _token_status()
    if info:
        return {"status": "ok", **info}
    raise HTTPException(status_code=401, detail="Token refresh failed")

@app.get("/v1/models")
async def list_models():
    return ModelList(data=[ModelInfo(id=n, created=int(time.time())) for n in MODEL_TONES])

@app.post("/v1/chat/completions")
async def chat_completions(body: ChatCompletionRequest, request: Request):
    try:
        token = resolve_token()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=401, detail=str(e))

    if not body.messages:
        raise HTTPException(status_code=400, detail="messages required")

    # Build tool list
    tools = []
    if body.tools:
        for t in body.tools:
            if t.type == "function":
                tools.append({"type": "function", "function": {
                    "name": t.function.name, "description": t.function.description,
                    "parameters": t.function.parameters,
                }})
    else:
        tools = BUILTIN_TOOLS

    # Get last user message
    last = ""
    for m in reversed(body.messages):
        if m.role == "user":
            last = m.content
            break
    if not last:
        raise HTTPException(status_code=400, detail="no user message")

    # Inject tools into messages for the model context
    # (We don't send system prompt to M365 directly, but we inject along with user msg)
    if tools:
        last = f"[Tools available: {json.dumps(tools, ensure_ascii=False)}]\n\n{last}"

    session, conv = _session(token, body.messages)
    tone = MODEL_TONES.get(body.model, "magic")
    cid = _make_chat_id()

    logger.info("Chat: model=%s tone=%s tools=%d conv=%s", body.model, tone, len(tools), conv[:8])

    async def do_chat() -> dict:
        """Non-streaming: send message, check for tool calls, loop."""
        msg = last
        max_round = 6
        for r in range(max_round):
            resp = await session.send(msg, tone=tone, on_delta=None)
            if resp.error:
                return {"id": cid, "object": "chat.completion", "created": int(time.time()),
                        "model": body.model, "choices": [{"index": 0, "message": {"role": "assistant", "content": resp.error}, "finish_reason": "stop"}]}
            if resp.disengaged:
                return {"id": cid, "object": "chat.completion", "created": int(time.time()),
                        "model": body.model, "choices": [{"index": 0, "message": {"role": "assistant", "content": "(disengaged)"}, "finish_reason": "stop"}]}

            text = resp.text
            tcs = _find_tool_calls(text)

            # Strip tool call block from displayed text
            clean = re.sub(r'```tool_call.*?```\s*', '', text, flags=re.DOTALL).strip()
            # Remove any bare tool JSON
            clean = re.sub(r'\{"tool".*?"\}\}', '', clean).strip()

            if not tcs:
                # No more tool calls → final answer
                resp_text = clean or text
                return {"id": cid, "object": "chat.completion", "created": int(time.time()),
                        "model": body.model,
                        "choices": [{"index": 0, "message": {"role": "assistant", "content": resp_text}, "finish_reason": "stop"}],
                        "usage": {"prompt_tokens": len(msg), "completion_tokens": len(resp_text)}}

            # Execute tool calls
            results = []
            for tc in tcs:
                name = tc.get("tool", "?")
                args = tc.get("arguments", {})
                result = _execute_tool(name, args)
                results.append(f"[{name}]\n{result[:500]}")

            msg = "Tool results:\n" + "\n\n".join(results)

        return {"id": cid, "object": "chat.completion", "created": int(time.time()),
                "model": body.model, "choices": [{"index": 0, "message": {"role": "assistant", "content": "Max rounds reached"}, "finish_reason": "stop"}]}

    async def stream_chat() -> AsyncGenerator[str, None]:
        """Streaming: yield SSE chunks."""
        msg = last
        max_round = 6
        for r in range(max_round):
            resp = await session.send(msg, tone=tone, on_delta=None)
            if resp.error or resp.disengaged:
                yield _delta(body.model, cid, "")
                yield "data: [DONE]\n\n"
                return

            text = resp.text
            tcs = _find_tool_calls(text)
            clean = re.sub(r'```tool_call.*?```\s*', '', text, flags=re.DOTALL).strip()
            clean = re.sub(r'\{"tool".*?"\}\}', '', clean).strip()

            if not tcs:
                if clean:
                    yield _delta(body.model, cid, clean)
                yield _delta(body.model, cid, "", fr="stop")
                yield "data: [DONE]\n\n"
                return

            # Execute tools
            results = []
            for tc in tcs:
                name = tc.get("tool", "?")
                args = tc.get("arguments", {})
                result = _execute_tool(name, args)
                results.append(f"[{name}]\n{result[:500]}")

            msg = "Tool results:\n" + "\n\n".join(results)

        yield _delta(body.model, cid, "", fr="stop")
        yield "data: [DONE]\n\n"

    if body.stream:
        return StreamingResponse(stream_chat(), media_type="text/event-stream",
                                  headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

    return await do_chat()

@app.get("/health")
async def health():
    return {"status": "ok", "models": list(MODEL_TONES.keys())}

# --- Main ---

def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=23100)
    p.add_argument("-v", "--verbose", action="store_true")
    a = p.parse_args()
    if a.verbose:
        logging.getLogger("m365").setLevel(logging.DEBUG)
    try:
        resolve_token()
    except Exception as e:
        logger.warning("No token: %s", e)
    uvicorn.run("api:app", host=a.host, port=a.port)

if __name__ == "__main__":
    main()
