"""Abstract LLM client and shared types."""

from __future__ import annotations

import abc
import asyncio
import json
import logging
import re

from pydantic import BaseModel, Field

from excel_eval.models import LLMConfig

logger = logging.getLogger(__name__)


# ── Response model ─────────────────────────────────────────────────────────


class LLMResponse(BaseModel):
    """Structured response from an LLM call."""

    content: str = ""
    input_tokens: int = 0
    output_tokens: int = 0
    latency_ms: int = 0
    cost_estimate: float = 0.0


# ── Abstract client ────────────────────────────────────────────────────────


class BaseLLMClient(abc.ABC):
    """Provider-agnostic async LLM client."""

    def __init__(self, config: LLMConfig) -> None:
        self.config = config

    @abc.abstractmethod
    async def complete(
        self,
        messages: list[dict],
        images: list[bytes] | None = None,
        json_schema: dict | None = None,
    ) -> LLMResponse:
        """Send a chat completion request and return the response."""

    async def complete_with_retry(
        self,
        messages: list[dict],
        images: list[bytes] | None = None,
        max_retries: int | None = None,
        json_mode: bool = False,
        json_schema: dict | None = None,
    ) -> LLMResponse:
        """Call *complete* with exponential-backoff retry and optional JSON
        validation.

        Retry policy
        ------------
        * API errors (rate-limit, 5xx, timeout, connection) are retried up to
          *max_retries* times with exponential backoff (2 s, 8 s, 32 s, …).
        * When *json_schema* is provided, the provider uses constrained
          decoding (e.g. Anthropic ``output_config``) to guarantee valid JSON.
        * When *json_mode* is ``True`` (without schema), the raw content is
          validated as JSON. If parsing fails, up to 2 additional attempts are
          made with an extra user message asking for valid JSON.
        """
        if max_retries is None:
            max_retries = self.config.max_retries

        last_exc: BaseException | None = None

        for attempt in range(max_retries):
            try:
                response = await self.complete(messages, images, json_schema=json_schema)
                break
            except Exception as exc:
                last_exc = exc
                if not self._is_retryable(exc):
                    raise
                delay = 2 ** (2 * attempt + 1)  # 2, 8, 32, …
                logger.warning(
                    "LLM API error (attempt %d/%d), retrying in %ds: %s",
                    attempt + 1,
                    max_retries,
                    delay,
                    exc,
                )
                await asyncio.sleep(delay)
        else:
            raise last_exc  # type: ignore[misc]

        # ── JSON validation retries (only when no schema enforcement) ──
        if json_mode and not json_schema:
            json_retry_msg = {
                "role": "user",
                "content": (
                    "Your previous response was not valid JSON. "
                    "Please respond with valid JSON only."
                ),
            }
            for json_attempt in range(2):
                try:
                    json.loads(self._strip_markdown_fences(response.content))
                    break
                except (json.JSONDecodeError, ValueError):
                    logger.warning(
                        "Response is not valid JSON (json retry %d/2), "
                        "requesting correction.",
                        json_attempt + 1,
                    )
                    retry_messages = messages + [
                        {"role": "assistant", "content": response.content},
                        json_retry_msg,
                    ]
                    response = await self.complete(retry_messages, images)

        return response

    @staticmethod
    def _strip_markdown_fences(text: str) -> str:
        """Strip markdown code fences before JSON parsing."""
        text = text.strip()
        fence_match = re.search(
            r"```(?:json)?\s*\n?(.*?)\n?\s*```", text, re.DOTALL
        )
        if fence_match:
            return fence_match.group(1).strip()
        return text

    # ── Helpers ─────────────────────────────────────────────────────────

    @staticmethod
    def _is_retryable(exc: BaseException) -> bool:
        """Return ``True`` for transient errors worth retrying."""
        # Generic transient errors
        if isinstance(exc, (TimeoutError, ConnectionError, OSError)):
            return True
        # Provider-specific errors (check by attribute to avoid hard dependency)
        exc_type_name = type(exc).__name__
        if exc_type_name in (
            "RateLimitError",
            "InternalServerError",
            "APITimeoutError",
            "APIConnectionError",
        ):
            return True
        return False


# ── Factory ────────────────────────────────────────────────────────────────


def create_llm_client(config: LLMConfig) -> BaseLLMClient:
    """Instantiate the correct provider based on *config.provider*."""
    provider = config.provider.lower()

    if provider == "anthropic":
        from excel_eval.llm.providers.anthropic import AnthropicLLMClient

        return AnthropicLLMClient(config)

    raise ValueError(f"Unsupported LLM provider: {config.provider!r}")
