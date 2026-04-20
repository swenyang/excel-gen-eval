"""Scenario auto-detection via a lightweight LLM call."""

from __future__ import annotations

import json
import logging
import re
import time
from pathlib import Path

from excel_eval.llm.base import BaseLLMClient
from excel_eval.llm.schemas import SCENARIO_DETECT_SCHEMA
from excel_eval.models import PreparedData, Scenario, ScenarioResult

logger = logging.getLogger(__name__)

_PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"
_MAX_ROWS_PER_SHEET = 5
_MAX_CONTENT_CHARS = 500


class ScenarioDetector:
    """Detects the scenario type of an Excel workbook via a lightweight LLM call."""

    def __init__(self, llm_client: BaseLLMClient) -> None:
        self.llm_client = llm_client

    async def detect(self, data: PreparedData) -> ScenarioResult:
        """Detect the scenario from user prompt + sheet names + content summary."""
        start = time.perf_counter()
        try:
            system_prompt = self._load_prompt()
            context = self._build_context(data)

            messages: list[dict] = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": context},
            ]

            response = await self.llm_client.complete_with_retry(
                messages, json_mode=True,
                json_schema=SCENARIO_DETECT_SCHEMA,
            )

            parsed = self._parse_response(response.content)

            scenario_str = parsed.get("scenario", "general")
            confidence = float(parsed.get("confidence", 0.0))
            reasoning = parsed.get("reasoning", "")
            raw_blend = parsed.get("blend", {})

            # Fall back to general when confidence is low
            try:
                detected = Scenario(scenario_str)
            except ValueError:
                detected = Scenario.GENERAL
                confidence = 0.0

            if confidence < 0.7:
                detected = Scenario.GENERAL

            # Validate and normalize blend
            blend: dict[str, float] = {}
            if raw_blend and isinstance(raw_blend, dict):
                total = sum(float(v) for v in raw_blend.values() if isinstance(v, (int, float)))
                for k, v in raw_blend.items():
                    try:
                        Scenario(k)  # validate scenario name
                        blend[k] = float(v) / total if total > 0 else 0.0
                    except (ValueError, TypeError):
                        pass
            if not blend:
                blend = {detected.value: 1.0}

            # Parse dimension applicability
            raw_applicable = parsed.get("applicable_dimensions", {})
            applicable_dimensions: dict[str, bool] = {}
            if raw_applicable and isinstance(raw_applicable, dict):
                for k, v in raw_applicable.items():
                    if isinstance(v, bool):
                        applicable_dimensions[k] = v
            dim_reasoning = parsed.get("dimension_reasoning", {})
            if not isinstance(dim_reasoning, dict):
                dim_reasoning = {}

            elapsed_ms = int((time.perf_counter() - start) * 1000)
            blend_str = ", ".join(f"{k}={v:.0%}" for k, v in blend.items())
            na_dims = [k for k, v in applicable_dimensions.items() if not v]
            na_str = f", n/a=[{', '.join(na_dims)}]" if na_dims else ""
            logger.info(
                "Scenario detected: %s (confidence=%.2f, blend=[%s]%s) in %dms",
                detected.value,
                confidence,
                blend_str,
                na_str,
                elapsed_ms,
            )
            return ScenarioResult(
                detected=detected,
                confidence=confidence,
                reasoning=reasoning,
                blend=blend,
                applicable_dimensions=applicable_dimensions,
                dimension_reasoning=dim_reasoning,
            )

        except Exception as exc:
            logger.exception("Scenario detection failed: %s", exc)
            return ScenarioResult(
                detected=Scenario.GENERAL,
                confidence=0.0,
                reasoning=f"Detection failed: {exc}",
            )

    # ── Helpers ────────────────────────────────────────────────────────

    @staticmethod
    def _build_context(data: PreparedData) -> str:
        """Build a compact context string for scenario classification."""
        parts: list[str] = []

        if data.user_prompt:
            parts.append(f"## User Prompt\n{data.user_prompt}")

        sheet_names = [s.name for s in data.visible_sheets]
        parts.append(f"## Sheet Names\n{', '.join(sheet_names) if sheet_names else '(none)'}")

        # Include a brief content summary (first few rows per sheet)
        summaries: list[str] = []
        for sheet in data.visible_sheets:
            lines = sheet.csv_text.splitlines()
            preview = "\n".join(lines[: _MAX_ROWS_PER_SHEET])
            if len(preview) > _MAX_CONTENT_CHARS:
                preview = preview[:_MAX_CONTENT_CHARS] + "…"
            summaries.append(
                f"### {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols)\n{preview}"
            )
        if summaries:
            parts.append("## Content Preview\n" + "\n\n".join(summaries))

        # Mention formulas/charts presence
        if data.formulas:
            parts.append(f"## Formulas\n{len(data.formulas)} formula(s) detected.")
        if data.charts:
            chart_types = [c.chart_type for c in data.charts]
            parts.append(f"## Charts\n{len(data.charts)} chart(s): {', '.join(chart_types)}")

        return "\n\n".join(parts)

    def _load_prompt(self) -> str:
        path = _PROMPTS_DIR / "scenario_detection.md"
        if not path.exists():
            raise FileNotFoundError(f"Prompt template not found: {path}")
        return path.read_text(encoding="utf-8")

    @staticmethod
    def _parse_response(content: str) -> dict:
        """Parse LLM JSON response, handling markdown code fences."""
        text = content.strip()
        fence_match = re.search(
            r"```(?:json)?\s*\n?(.*?)\n?\s*```", text, re.DOTALL
        )
        if fence_match:
            text = fence_match.group(1).strip()
        return json.loads(text)
