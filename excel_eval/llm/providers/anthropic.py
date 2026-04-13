"""Anthropic Claude provider implementation."""

from __future__ import annotations

import base64
import os
import time
from typing import Any

import anthropic

from excel_eval.llm.base import BaseLLMClient, LLMResponse
from excel_eval.models import LLMConfig

# Per-token costs (USD) keyed by model-name substring.
# Input / output prices sourced from public Anthropic pricing.
_MODEL_COSTS: dict[str, tuple[float, float]] = {
    "claude-sonnet-4": (3.0e-6, 15.0e-6),
    "claude-3-7-sonnet": (3.0e-6, 15.0e-6),
    "claude-3-5-sonnet": (3.0e-6, 15.0e-6),
    "claude-3-5-haiku": (0.8e-6, 4.0e-6),
    "claude-3-haiku": (0.25e-6, 1.25e-6),
    "claude-3-opus": (15.0e-6, 75.0e-6),
    "claude-opus-4": (15.0e-6, 75.0e-6),
}

# Fallback if model name doesn't match any known pattern.
_DEFAULT_COST: tuple[float, float] = (3.0e-6, 15.0e-6)


class AnthropicLLMClient(BaseLLMClient):
    """LLM client backed by the Anthropic Messages API."""

    def __init__(self, config: LLMConfig) -> None:
        super().__init__(config)
        api_key = config.api_key
        if api_key is None and config.api_key_env:
            api_key = os.environ.get(config.api_key_env)
        self._client = anthropic.AsyncAnthropic(
            api_key=api_key,
            timeout=anthropic.Timeout(
                timeout=config.timeout,      # total timeout (default 120s)
                connect=30.0,                # connect timeout
                read=config.timeout,         # per-read timeout (prevents streaming hangs)
            ),
        )

    # ── Public API ─────────────────────────────────────────────────────

    async def complete(
        self,
        messages: list[dict],
        images: list[bytes] | None = None,
        json_schema: dict | None = None,
    ) -> LLMResponse:
        system_text, api_messages = self._convert_messages(messages, images)

        kwargs: dict[str, Any] = {
            "model": self.config.model,
            "max_tokens": self.config.max_tokens,
            "temperature": self.config.temperature,
            "messages": api_messages,
        }
        if system_text:
            kwargs["system"] = system_text

        if json_schema:
            kwargs["output_config"] = {
                "format": {
                    "type": "json_schema",
                    "schema": json_schema,
                }
            }
        t0 = time.perf_counter()
        response = await self._client.messages.create(**kwargs)
        latency_ms = int((time.perf_counter() - t0) * 1000)

        content = self._extract_text(response)
        input_tokens = response.usage.input_tokens
        output_tokens = response.usage.output_tokens
        cost = self._estimate_cost(input_tokens, output_tokens)

        return LLMResponse(
            content=content,
            input_tokens=input_tokens,
            output_tokens=output_tokens,
            latency_ms=latency_ms,
            cost_estimate=cost,
        )

    # ── Message conversion ─────────────────────────────────────────────

    @staticmethod
    def _convert_messages(
        messages: list[dict],
        images: list[bytes] | None = None,
    ) -> tuple[str, list[dict]]:
        """Convert generic chat messages to Anthropic format.

        Returns ``(system_text, api_messages)``.
        """
        system_parts: list[str] = []
        api_messages: list[dict] = []

        for msg in messages:
            role = msg.get("role", "user")
            text = msg.get("content", "")

            if role == "system":
                system_parts.append(text)
                continue

            api_messages.append({"role": role, "content": text})

        # Attach images to the last user message as content blocks.
        if images and api_messages:
            last_user_idx = _last_index(api_messages, "user")
            if last_user_idx is not None:
                existing_text = api_messages[last_user_idx]["content"]
                content_blocks: list[dict] = []

                for img_bytes in images:
                    media_type = _detect_media_type(img_bytes)
                    content_blocks.append(
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": base64.b64encode(img_bytes).decode(),
                            },
                        }
                    )

                # Append the original text after the images.
                if existing_text:
                    content_blocks.append({"type": "text", "text": existing_text})

                api_messages[last_user_idx]["content"] = content_blocks

        system_text = "\n\n".join(system_parts)
        return system_text, api_messages

    # ── Helpers ─────────────────────────────────────────────────────────

    @staticmethod
    def _extract_text(response: Any) -> str:
        parts: list[str] = []
        for block in response.content:
            if hasattr(block, "text"):
                parts.append(block.text)
        return "".join(parts)

    def _estimate_cost(self, input_tokens: int, output_tokens: int) -> float:
        model = self.config.model.lower()
        for key, (in_cost, out_cost) in _MODEL_COSTS.items():
            if key in model:
                return input_tokens * in_cost + output_tokens * out_cost
        in_cost, out_cost = _DEFAULT_COST
        return input_tokens * in_cost + output_tokens * out_cost


# ── Module-level helpers ───────────────────────────────────────────────────


def _last_index(messages: list[dict], role: str) -> int | None:
    for i in range(len(messages) - 1, -1, -1):
        if messages[i].get("role") == role:
            return i
    return None


def _detect_media_type(data: bytes) -> str:
    if data[:8].startswith(b"\x89PNG"):
        return "image/png"
    if data[:3] == b"\xff\xd8\xff":
        return "image/jpeg"
    if data[:4] == b"GIF8":
        return "image/gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    return "image/png"  # safe default
