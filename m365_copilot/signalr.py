"""
M365 Copilot SignalR Client — raw WebSocket + SignalR JSON protocol.

Handles:
- WebSocket connection with browser-like headers
- SignalR JSON framing (0x1E record separator)
- Handshake, ping/pong, chat invocation, metrics frame
- Streaming response parsing (delta + full message updates)
- Disengaged detection, throttling info, scores extraction

Protocol spec: https://github.com/cramt/m365-copilot-proxy/blob/main/docs/m365-copilot-api.md
"""

import json
import logging
import time
import uuid
from typing import Callable, Optional

logger = logging.getLogger("m365.signalr")

RS = "\x1E"  # Record Separator (0x1E) — SignalR JSON frame terminator

# Feature variants cargo-culted from m365.cloud.microsoft captured session
VARIANTS = ",".join([
    "EnableMcpServerWidgets",
    "feature.EnableMcpServerWidgets",
    "feature.EnableLuForChatCIQ",
    "feature.enableChatCIQPlugin",
    "EnableRequestPlugins",
    "feature.EnableSensitivityLabels",
    "feature.IsCustomEngineCopilotEnabled",
    "feature.bizchatfluxv3",
    "feature.enablechatpages",
    "feature.enableCodeCanvas",
    "feature.turnOnWorkTabRecommendation",
    "turnOffWorkTabUpsellFromClient",
    "feature.turnOnDARecommendation",
    "feature.IsStreamingModeInChatRequestEnabled",
    "IncludeSourceAttributionsConcise",
    "SkipPublishEmptyMessage",
    "feature.EnableDeduplicatingSourceAttributions",
    "Enable3PActionProgressMessages",
    "feature.enableClientWebRtc",
    "feature.EnableMeetingRecapOfSeriesMeetingWithCiq",
    "feature.EnableReferencesListCompleteSignal",
    "feature.StorageMessageSplitDisabled",
    "feature.EnableCuaTakeControlApi",
    "feature.cwcallowedos",
    "feature.disabledisallowedmsgs",
    "feature.enableCitationsForSynthesisData",
    "feature.enableGenerateGraphicArtOptionsSet",
    "cdximagen",
    "feature.EnableUpdatedUXForConfirmationDialog",
    "feature.EnableClientFileURLSupportForOfficeWebPaidCopilot",
    "feature.EnableDesignEditorImageGrounding",
    "feature.EnableDesignerEditor",
    "feature.OfficeWebToHelix",
    "feature.OfficeDesktopToHelix",
    "feature.M365TeamsHubToHelix",
    "feature.OwaHubToHelix",
    "feature.MonarchHubToHelix",
    "feature.Win32OutlookHubToHelix",
    "feature.MacOutlookHubToHelix",
    "Agt_bizchat_enableGpt5ForHelix",
])

# Code interpreter options — enables real server-side Python execution
CODE_INTERPRETER_OPTIONS = [
    "cwc_code_interpreter",
    "cwc_code_interpreter_amsfix",
    "cwc_code_interpreter_citation_fix",
    "code_interpreter_interactive_charts",
    "code_interpreter_matplotlib_patching",
]

ALLOWED_MESSAGE_TYPES = [
    "Chat", "Suggestion", "Progress", "EndOfRequest",
    "GeneratedCode", "GenerateContentQuery", "ReferencesListComplete",
    "RenderCardRequest", "Disengaged", "InternalSearchQuery",
    "InternalSearchResult", "InternalSearchResultPreview",
]


class ChatResponse:
    """Parsed response from a chat turn."""

    def __init__(self):
        self.delta_text = ""
        self.full_text = ""
        self.has_content = False
        self.message_type: Optional[str] = None
        self.throttle: Optional[dict] = None
        self.content_origin: Optional[str] = None
        self.turn_count: Optional[int] = None
        self.turn_state: Optional[str] = None
        self.scores: dict[str, float] = {}
        self.disengaged = False
        self.error: Optional[str] = None
        self.raw_frames: list[dict] = []
        self._last_full_text_len = 0  # track what we've already shown

    @property
    def text(self) -> str:
        """Best-effort accumulated text. Prefers full message snapshots."""
        return self.full_text if len(self.full_text) >= len(self.delta_text) else self.delta_text

    @property
    def new_text(self) -> str:
        """Text that hasn't been returned via on_delta yet."""
        current = self.text
        new_part = current[self._last_full_text_len:]
        if new_part:
            self._last_full_text_len = len(current)
        return new_part

    def __repr__(self):
        status = "disengaged" if self.disengaged else ("ok" if self.has_content else "empty")
        return f"<ChatResponse {status} text={self.text[:100]!r}>"


def build_ws_url(token: str, conversation_id: str, session_id: str, request_id: Optional[str] = None) -> str:
    """Build the WebSocket URL for the M365 Copilot SignalR hub."""
    from .auth import decode_jwt

    claims = decode_jwt(token)
    oid = claims["oid"]
    tid = claims["tid"]

    rid = request_id or str(uuid.uuid4())
    import urllib.parse
    params = {
        "chatsessionid": rid,
        "clientrequestid": rid,
        "X-SessionId": session_id,
        "ConversationId": conversation_id,
        "access_token": token,
        "variants": VARIANTS,
        "source": '"officeweb"',
        "product": "Office",
        "agentHost": "Bizchat.FullScreen",
        "licenseType": "Starter",
        "agent": "web",
        "scenario": "OfficeWebIncludedCopilot",
    }
    query = urllib.parse.urlencode(params, doseq=True)
    return f"wss://substrate.office.com/m365Copilot/Chathub/{oid}@{tid}?{query}"


def build_chat_frame(
    text: str,
    tone: str = "magic",
    is_first: bool = True,
    agent_id: Optional[str] = None,
    enable_code_interpreter: bool = False,
) -> str:
    """Build the SignalR chat invocation frame (type:4, target:chat)."""
    options_sets = []
    if enable_code_interpreter:
        options_sets.extend(CODE_INTERPRETER_OPTIONS)

    body = {
        "message": {
            "text": text,
            "author": "user",
        },
        "tone": tone,
        "source": "officeweb",
        "streamingMode": "ConciseWithPadding",
        "isStartOfSession": is_first,
        "allowedMessageTypes": ALLOWED_MESSAGE_TYPES,
        "optionsSets": options_sets,
        "clientInfo": {
            "clientPlatform": "mcmcopilot-web",
            "clientAppName": "Office",
            "clientAppVersion": "1.0.0",
        },
        "plugins": [{"Id": "BingWebSearch", "Source": "BuiltIn"}] if is_first else [],
        "locationInfo": {
            "timeZone": "Asia/Shanghai",
            "locale": "zh-CN",
        },
    }

    if agent_id:
        body["threadLevelGptId"] = {"id": agent_id, "source": "MOS3"}
        body["gpts"] = [{
            "id": agent_id,
            "source": "MOS3",
            "version": "1.0.0",
            "clientOverrides": {
                "capabilities": [],
                "deepResearchModels@odata.type": "Collection(String)",
            },
        }]
        body.pop("plugins", None)

    frame = {
        "type": 4,
        "invocationId": "0",
        "target": "chat",
        "arguments": [body],
    }
    return json.dumps(frame) + RS


def build_metrics_frame() -> str:
    """Build the mandatory Metrics frame sent alongside the chat invocation."""
    now = time.strftime("%Y-%m-%dT%H:%M:%S.000Z", time.gmtime())
    metrics = {
        "type": 1,
        "target": "Metrics",
        "arguments": [{
            "Timestamps": {
                "ConnectionStart": now,
                "UserInputStart": now,
                "ConnectionEstablished": now,
                "UserInputSubmit": now,
            },
        }],
    }
    return json.dumps(metrics) + RS


def _process_update_frame(raw: dict, response: ChatResponse) -> Optional[str]:
    """Process a type:1 update frame. Returns any newly available text."""
    args = raw.get("arguments", [])

    for arg in args:
        if not isinstance(arg, dict):
            continue

        # Full message snapshot — update full_text
        if "messages" in arg:
            for msg in arg["messages"]:
                if not isinstance(msg, dict):
                    continue

                msg_type = msg.get("messageType")
                response.message_type = msg_type

                # Disengaged detection
                if msg_type == "Disengaged":
                    response.disengaged = True
                    continue

                # Real bot content — update full_text
                if not msg_type and msg.get("author") == "bot":
                    t = msg.get("text", "")
                    if t.strip():
                        response.full_text = t
                        response.has_content = True

                # Scores
                scores = msg.get("scores")
                if scores:
                    for s in scores:
                        if isinstance(s, dict):
                            comp = s.get("component", "")
                            score = s.get("score", 0)
                            response.scores[comp] = max(response.scores.get(comp, 0), score)

                # Other metadata
                for field in ["contentOrigin", "turnCount", "turnState"]:
                    val = msg.get(field)
                    if val is not None and isinstance(val, (str, int)):
                        setattr(response, field, val)

        # Delta update — accumulate text
        if "writeAtCursor" in arg:
            text = arg["writeAtCursor"]
            response.delta_text += text
            if text.strip():
                response.has_content = True

        # Throttling info
        if "throttling" in arg:
            t = arg["throttling"]
            if isinstance(t, dict):
                response.throttle = {
                    "current": t.get("numUserMessagesInConversation", 0),
                    "max": t.get("maxNumUserMessagesInConversation", 600),
                }

    # Return any new text since last check
    return response.new_text or None


async def chat(
    token: str,
    text: str,
    tone: str = "magic",
    agent_id: Optional[str] = None,
    conversation_id: Optional[str] = None,
    session_id: Optional[str] = None,
    is_first: bool = True,
    enable_code_interpreter: bool = False,
    on_delta: Optional[Callable[[str], None]] = None,
) -> ChatResponse:
    """
    Send a chat message and stream the response from M365 Copilot.

    Opens a fresh WebSocket per turn (as the official client does).
    Reuse conversation_id and session_id across turns for server-side context.

    Args:
        token: Sydney JWT access token
        text: The user message
        tone: Model selector ('magic', 'Claude_Sonnet', 'Gpt_5_4_Reasoning', etc.)
        agent_id: Optional Copilot Studio agent ID for tool calling
        conversation_id: Reuse to continue a conversation
        session_id: Reuse across turns for server-side context
        is_first: True for the first turn in a conversation
        enable_code_interpreter: Enable server-side Python execution
        on_delta: Optional callback for streaming text chunks

    Returns:
        ChatResponse with accumulated text and metadata
    """
    from .auth import decode_jwt

    conv_id = conversation_id or str(uuid.uuid4())
    sess_id = session_id or str(uuid.uuid4())
    req_id = str(uuid.uuid4())
    ws_url = build_ws_url(token, conv_id, sess_id, request_id=req_id)
    response = ChatResponse()

    # Decode token for claims info
    claims = decode_jwt(token)
    oid = claims["oid"]
    logger.debug("Connecting: oid=%s, conv=%s", oid[:8], conv_id[:8])

    try:
        import websockets.asyncio.client as ws_client
    except ImportError:
        response.error = "websockets package not installed. Run: pip install websockets"
        return response

    try:
        async with ws_client.connect(
            ws_url,
            additional_headers={
                "Origin": "https://m365.cloud.microsoft",
                "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:148.0) Gecko/20100101 Firefox/148.0",
                "Accept-Language": "en-US,en;q=0.9",
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
            },
            close_timeout=30,
            max_size=10 * 1024 * 1024,
        ) as ws:
            # Step 1: SignalR handshake
            handshake = json.dumps({"protocol": "json", "version": 1}) + RS
            await ws.send(handshake)

            # Wait for handshake ACK
            handshake_resp = await ws.recv()
            if isinstance(handshake_resp, bytes):
                handshake_resp = handshake_resp.decode("utf-8")
            logger.debug("Handshake: %s", handshake_resp[:80])

            # Step 2: Send chat + Metrics frame together
            payload = build_chat_frame(text, tone, is_first, agent_id, enable_code_interpreter)
            payload += build_metrics_frame()
            await ws.send(payload)

            # Step 3: Read response frames
            async for message in ws:
                if isinstance(message, bytes):
                    message = message.decode("utf-8", errors="replace")

                # SignalR: multiple 0x1E-separated frames in one message
                for chunk in message.split(RS):
                    chunk = chunk.strip()
                    if not chunk:
                        continue

                    try:
                        raw = json.loads(chunk)
                    except json.JSONDecodeError:
                        logger.warning("Bad JSON: %s", chunk[:200])
                        continue

                    response.raw_frames.append(raw)
                    frame_type = raw.get("type")

                    if frame_type == 1:
                        logger.debug("Frame type=1 target=%s args_keys=%s",
                            raw.get("target"), list(raw.get("arguments", [{}])[0].keys()) if raw.get("arguments") else "no args")
                    elif frame_type == 2:
                        logger.debug("Frame type=2 keys=%s", list(raw.keys()))

                    # Type 1: Update (streaming text, messages, throttling)
                    if frame_type == 1:
                        delta = _process_update_frame(raw, response)
                        if delta is not None:
                            if on_delta:
                                on_delta(delta)

                    # Type 2: Stream item (end of turn with full state)
                    elif frame_type == 2:
                        item = raw.get("item", {}) or {}
                        result_data = item.get("result", {})
                        result_val = result_data.get("value")

                        logger.debug("Type 2 item keys: %s", list(item.keys()))
                        logger.debug("Type 2 result: %s", result_data)

                        if result_val and result_val != "Success":
                            response.error = f"Server returned result.value={result_val}"
                            if result_data.get("message"):
                                response.error += f": {result_data['message']}"
                            response.has_content = False
                            return response

                        # Check final messages
                        for msg in item.get("messages", []):
                            if isinstance(msg, dict) and msg.get("author") == "bot":
                                if not msg.get("messageType"):
                                    t = msg.get("text", "")
                                    if t.strip():
                                        response.full_text = t
                                        response.has_content = True

                        # Throttling from item
                        throttling = item.get("throttling", {})
                        if throttling:
                            response.throttle = {
                                "current": throttling.get("numUserMessagesInConversation", 0),
                                "max": throttling.get("maxNumUserMessagesInConversation", 600),
                            }

                        # Turn state / count
                        ts = item.get("turnState")
                        if ts:
                            response.turn_state = str(ts)
                        tc = item.get("telemetry", {}).get("turnCount")
                        if tc is not None:
                            response.turn_count = int(tc)

                        return response  # End of turn

                    # Type 3: Completion
                    elif frame_type == 3:
                        err = raw.get("error") or (raw.get("result") or {}).get("message")
                        if err:
                            response.error = str(err)
                        return response

                    # Type 6: Ping — reply with pong
                    elif frame_type == 6:
                        await ws.send(json.dumps({"type": 6}) + RS)

                    # Type 7: Close
                    elif frame_type == 7:
                        err = raw.get("error")
                        if err:
                            response.error = str(err)
                        return response

    except Exception as e:
        response.error = f"Connection error: {e}"
        logger.exception("Chat error")

    return response
