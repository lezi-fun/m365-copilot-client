"""
M365 Copilot Session — maintains conversation state across turns.
"""

import uuid
import logging
from typing import Optional

from .signalr import chat, ChatResponse, ALLOWED_MESSAGE_TYPES

logger = logging.getLogger("m365.session")


# Model name → tone mapping (from m365-copilot-proxy's MODEL_TONES)
MODEL_TONES = {
    "auto": "magic",
    "quick": "Gpt_Quick",
    "think-deeper": "Gpt_Reasoning",
    "claude": "Claude_Sonnet",
    "claude-sonnet": "Claude_Sonnet",
    "claude-sonnet-4.5": "Claude_Sonnet",
    "claude-sonnet-think-deeper": "Claude_Sonnet_Reasoning",
    "claude-opus": "Claude_Opus",
    "gpt-5.5": "Gpt_5_5_Chat",
    "gpt-5.5-quick": "Gpt_5_5_Chat",
    "gpt-5.5-think-deeper": "Gpt_5_5_Reasoning",
    "gpt-5.4": "Gpt_5_4_Reasoning",
    "gpt-5.4-quick": "Gpt_5_4_Quick",
    "gpt-5.4-think-deeper": "Gpt_5_4_Reasoning",
    "gpt-5.3": "Gpt_5_3_Quick",
    "gpt-5.3-quick": "Gpt_5_3_Quick",
    "gpt-5.3-think-deeper": "Gpt_5_3_Reasoning",
    "gpt-5.2": "Gpt_5_2_Quick",
    "gpt-5.2-quick": "Gpt_5_2_Quick",
    "gpt-5.2-think-deeper": "Gpt_5_2_Reasoning",
}

DEFAULT_TONE = "magic"


class CopilotSession:
    """
    A persistent conversation with M365 Copilot.
    Reuses conversation_id/session_id across turns for server-side context.
    """

    def __init__(
        self,
        token: str,
        agent_id: Optional[str] = None,
        tone: str = DEFAULT_TONE,
        enable_code_interpreter: bool = False,
    ):
        self.token = token
        self.agent_id = agent_id
        self.tone = tone
        self.enable_code_interpreter = enable_code_interpreter
        self.conversation_id = str(uuid.uuid4())
        self.session_id = str(uuid.uuid4())
        self.turn_count = 0
        self.last_response: Optional[ChatResponse] = None

    def set_tone(self, model_name: str):
        """Set tone by model name."""
        tone = MODEL_TONES.get(model_name)
        if tone:
            self.tone = tone
        else:
            logger.warning("Unknown model: %s, using auto", model_name)

    @property
    def is_first_turn(self) -> bool:
        return self.turn_count == 0

    async def send(
        self,
        text: str,
        tone: Optional[str] = None,
        on_delta=None,
    ) -> ChatResponse:
        """
        Send a message in this conversation.

        Args:
            text: The user message
            tone: Override the session tone for this turn
            on_delta: Optional callback for streaming text

        Returns:
            ChatResponse with the bot's reply
        """
        turn_tone = tone or self.tone

        resp = await chat(
            token=self.token,
            text=text,
            tone=turn_tone,
            agent_id=self.agent_id,
            conversation_id=self.conversation_id,
            session_id=self.session_id,
            is_first=self.is_first_turn,
            enable_code_interpreter=self.enable_code_interpreter,
            on_delta=on_delta,
        )

        self.turn_count += 1
        self.last_response = resp
        return resp

    def reset(self):
        """Start a new conversation (new conversation_id)."""
        self.conversation_id = str(uuid.uuid4())
        self.turn_count = 0
        self.last_response = None
