"""Base class for all dimension evaluators."""

from __future__ import annotations

import abc
import io
import json
import logging
import re
import time
from pathlib import Path

from PIL import Image

from excel_eval.llm.base import BaseLLMClient, LLMResponse
from excel_eval.llm.schemas import DIMENSION_EVAL_SCHEMA, EVIDENCE_VERIFY_SCHEMA
from excel_eval.models import (
    DimensionName,
    DimensionResult,
    EvalStatus,
    PreparedData,
    Scenario,
)

logger = logging.getLogger(__name__)

_PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"


def format_sheet_header(sheet, extra: str = "") -> str:
    """Format a sheet header with visibility annotation."""
    from excel_eval.models import SheetData
    hidden_tag = " [HIDDEN SHEET]" if getattr(sheet, "hidden", False) else ""
    base = f"### Sheet: {sheet.name}{hidden_tag} ({sheet.row_count} rows × {sheet.col_count} cols)"
    if extra:
        base += f" {extra}"
    return base


def _is_feedback_incomplete(feedback: str) -> bool:
    """Check if feedback text appears to be cut off mid-sentence.

    Uses a whitelist of valid sentence-ending characters. If the feedback
    doesn't end with one of these, it's likely incomplete.
    """
    text = feedback.rstrip()
    if not text:
        return True
    last_char = text[-1]
    # Valid sentence endings (Chinese + English + closing brackets)
    complete_endings = set("。！？.!?）)」】》}]\"'")
    return last_char not in complete_endings


def _downscale_image(img_bytes: bytes, max_width: int = 1200) -> bytes:
    """Downscale a PNG image to fit within max_width, preserving aspect ratio.

    Default 1200px keeps text and numbers readable (including thousand
    separators) while staying within ~2-4 Anthropic tiles per image.
    Returns original bytes if already small enough.
    """
    img = Image.open(io.BytesIO(img_bytes))
    w, h = img.size
    if w <= max_width:
        return img_bytes
    scale = max_width / w
    new_w, new_h = int(w * scale), int(h * scale)
    resized = img.resize((new_w, new_h), Image.LANCZOS)
    buf = io.BytesIO()
    resized.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


class BaseEvaluator(abc.ABC):
    """Base class for all dimension evaluators."""

    def __init__(self, llm_client: BaseLLMClient, language: str = "zh") -> None:
        self.llm_client = llm_client
        self.language = language
        self._prompt_template: str | None = None

    # ── Abstract interface ─────────────────────────────────────────────

    @property
    @abc.abstractmethod
    def dimension(self) -> DimensionName:
        """The dimension this evaluator scores."""

    @property
    @abc.abstractmethod
    def prompt_file(self) -> str:
        """Filename of the prompt template (e.g., 'data_accuracy.md')."""

    @abc.abstractmethod
    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        """Build the dimension-specific context string from prepared data.

        This is where each evaluator selects which parts of *PreparedData* to
        include in the LLM prompt.
        """

    # ── Optional overrides ─────────────────────────────────────────────

    def needs_screenshots(self) -> bool:
        """Override to ``True`` for visual dimensions."""
        return False

    # ── Core evaluation logic ──────────────────────────────────────────

    async def evaluate(
        self, data: PreparedData, scenario: Scenario
    ) -> DimensionResult:
        """Run evaluation for this dimension.

        Handles prompt loading, LLM call, response parsing.  On any
        unrecoverable error the result is returned with
        ``status=EvalStatus.ERROR``.
        """
        start = time.perf_counter()
        try:
            system_prompt = self._load_prompt()

            # Always evaluate in English to avoid JSON truncation with CJK languages.
            # Non-English feedback is generated via post-evaluation translation.

            context = self.build_context(data, scenario)

            # Inject scenario info into the user context
            scenario_header = (
                f"## Scenario\n"
                f"- **Type**: {scenario.value}\n"
                f"- **User prompt**: {data.user_prompt}\n\n"
            )
            user_content = scenario_header + context

            messages: list[dict] = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content},
            ]

            # Collect images for visual evaluators — all sheets, downscaled for token efficiency
            images: list[bytes] | None = None
            screenshot_note = ""
            if self.needs_screenshots():
                if data.screenshots:
                    images = [
                        _downscale_image(img_bytes, max_width=800)
                        for img_bytes in data.screenshots.values()
                    ]
                else:
                    screenshot_note = (
                        "\n\n**Note:** Screenshots were requested but are "
                        "unavailable. Base your assessment only on the "
                        "textual data provided. Because visual inspection "
                        "is not possible, the maximum score for this "
                        "dimension is capped at 3.\n"
                    )
                    messages[-1]["content"] += screenshot_note

            response: LLMResponse = await self.llm_client.complete_with_retry(
                messages, images=images, json_mode=True,
                json_schema=DIMENSION_EVAL_SCHEMA,
            )

            # Guard 1: output truncated by max_tokens — JSON is likely broken
            if response.stop_reason == "max_tokens":
                # Check if the response is still valid JSON despite truncation
                try:
                    json.loads(response.content)
                except (json.JSONDecodeError, ValueError):
                    logger.warning(
                        "Dimension %s: output hit max_tokens (%d) and JSON is invalid, marking ERROR",
                        self.dimension, response.output_tokens,
                    )
                    elapsed_ms = int((time.perf_counter() - start) * 1000)
                    return DimensionResult(
                        dimension=self.dimension,
                        status=EvalStatus.ERROR,
                        error_message=(
                            f"LLM output truncated at max_tokens ({response.output_tokens}). "
                            f"Increase max_tokens in config or simplify the evaluation context."
                        ),
                        input_tokens=response.input_tokens,
                        output_tokens=response.output_tokens,
                        latency_ms=elapsed_ms,
                    )

            parsed = self._parse_response(response.content)
            score = parsed.get("score")
            feedback = parsed.get("feedback", "")

            # Guard 2: incomplete or short feedback — retry with nudge
            # Detects: (a) feedback < 100 chars, (b) feedback ends mid-sentence
            needs_retry = False
            if score is not None and response.stop_reason != "max_tokens":
                if len(feedback) < 100:
                    needs_retry = True
                elif _is_feedback_incomplete(feedback):
                    needs_retry = True

            if needs_retry:
                logger.warning(
                    "Dimension %s: incomplete feedback (%d chars, %d output tokens), retrying with nudge",
                    self.dimension, len(feedback), response.output_tokens,
                )
                retry_messages = list(messages)
                retry_messages.append({
                    "role": "assistant",
                    "content": response.content,
                })
                retry_messages.append({
                    "role": "user",
                    "content": (
                        "Your feedback appears incomplete (cut off mid-sentence). "
                        "Please provide the COMPLETE analysis with proper sentence endings "
                        "and at least 3 specific findings in the evidence array. "
                        "Respond with the complete JSON again."
                    ),
                })
                response = await self.llm_client.complete_with_retry(
                    retry_messages, images=images, json_mode=True,
                    json_schema=DIMENSION_EVAL_SCHEMA,
                )
                retry_parsed = self._parse_response(response.content)
                retry_feedback = retry_parsed.get("feedback", "")
                if len(retry_feedback) > len(feedback):
                    parsed = retry_parsed
                    score = parsed.get("score")
                    feedback = retry_feedback

            # Cap score when screenshots were needed but missing
            if (
                self.needs_screenshots()
                and not data.screenshots
                and score is not None
                and score > 3
            ):
                score = 3

            # Verify VERIFIED evidence items via independent LLM call
            evidence = parsed.get("evidence", [])
            verified_items = [e for e in evidence if e.strip().startswith("+") or e.strip().startswith("-")]
            if verified_items and "VERIFIED" in " ".join(verified_items).upper():
                evidence = await self._verify_evidence(evidence, data)

            elapsed_ms = int((time.perf_counter() - start) * 1000)
            total_input = response.input_tokens
            total_output = response.output_tokens
            total_cost = response.cost_estimate
            return DimensionResult(
                dimension=self.dimension,
                status=EvalStatus.SUCCESS,
                score=score,
                feedback=feedback,
                evidence=evidence,
                input_tokens=total_input,
                output_tokens=total_output,
                latency_ms=elapsed_ms,
                cost_estimate=total_cost,
            )

        except Exception as exc:
            elapsed_ms = int((time.perf_counter() - start) * 1000)
            logger.exception(
                "Evaluation failed for dimension %s: %s", self.dimension, exc
            )
            return DimensionResult(
                dimension=self.dimension,
                status=EvalStatus.ERROR,
                error_message=str(exc),
                latency_ms=elapsed_ms,
            )

    # ── Helpers ────────────────────────────────────────────────────────

    async def _verify_evidence(
        self, evidence: list[str], data: PreparedData,
    ) -> list[str]:
        """Independently verify VERIFIED evidence items via a second LLM call.

        Items that cannot be confirmed are downgraded from VERIFIED to UNCONFIRMED.
        INFERRED items pass through unchanged.
        """
        # Only verify items tagged VERIFIED
        to_verify = []
        for i, e in enumerate(evidence):
            if "VERIFIED" in e.upper() and not "INFERRED" in e.upper():
                to_verify.append((i, e))

        if not to_verify:
            return evidence

        # Build verification context — same data the evaluator saw
        context_parts = []
        if data.scan_report_text:
            context_parts.append(data.scan_report_text)

        # Source data (same smart sizing as evaluators)
        if data.grounding_data:
            grounding_lines = data.grounding_data.split("\n")
            if len(grounding_lines) > 100:
                sample = "\n".join(grounding_lines[:50] + [f"[... {len(grounding_lines)-80} lines omitted ...]"] + grounding_lines[-30:])
                context_parts.append(f"## Source Data (sampled)\n{sample}")
            else:
                context_parts.append(f"## Source Data\n{data.grounding_data}")

        # Generated data (same smart sizing)
        for sheet in data.visible_sheets:
            lines = sheet.csv_text.split("\n")
            if len(lines) > 60:
                preview = "\n".join(lines[:30] + [f"[... {len(lines)-45} rows omitted ...]"] + lines[-15:])
            else:
                preview = sheet.csv_text
            context_parts.append(f"{format_sheet_header(sheet)}\n{preview}")

        # Screenshots context note
        if data.screenshots:
            context_parts.append(f"\n*Note: {len(data.screenshots)} screenshot(s) were provided to the evaluator for visual assessment. Claims about visual formatting may be based on screenshots not available here — give benefit of the doubt for visual claims.*")

        numbered = "\n".join(
            f"{i+1}. {e}" for i, (_, e) in enumerate(to_verify)
        )

        context_text = "\n\n".join(context_parts)
        verify_prompt = (
            "You are a fact-checker for an Excel evaluation. Below are evidence claims. "
            "For each VERIFIED claim, determine if it can be confirmed or contradicted by the data.\n\n"
            "Rules:\n"
            "- If the data SUPPORTS the claim → confirmed: true\n"
            "- If the data CONTRADICTS the claim → confirmed: false\n"
            "- If the data is INSUFFICIENT to confirm or deny (e.g., the relevant rows were omitted) → confirmed: true (benefit of the doubt)\n"
            "- Claims referencing the scan report should be checked against the scan report\n"
            "- Claims about visual formatting based on screenshots: confirmed: true (you cannot see screenshots)\n\n"
            "Respond with valid JSON only: a list of objects, one per claim:\n"
            '```json\n[{"index": 1, "confirmed": true/false, "reason": "brief explanation"}]\n```\n\n'
            f"## Data Reference\n\n{context_text}\n\n"
            f"## Claims to Verify\n{numbered}"
        )

        try:
            response = await self.llm_client.complete_with_retry(
                [{"role": "user", "content": verify_prompt}],
                json_mode=True,
                json_schema=EVIDENCE_VERIFY_SCHEMA,
            )
            parsed = json.loads(response.content)
            # Support both wrapped {results: [...]} and raw [...] formats
            if isinstance(parsed, list):
                results = parsed
            else:
                results = parsed.get("results", [])

            # Downgrade unconfirmed items
            updated = list(evidence)
            for result in results:
                idx = result.get("index", 0) - 1
                if 0 <= idx < len(to_verify):
                    orig_idx, orig_text = to_verify[idx]
                    if not result.get("confirmed", True):
                        updated[orig_idx] = orig_text.replace(
                            "VERIFIED", "UNCONFIRMED"
                        )
                        logger.info(
                            "Evidence downgraded to UNCONFIRMED: %s (reason: %s)",
                            orig_text[:80],
                            result.get("reason", ""),
                        )
            return updated

        except Exception as exc:
            logger.warning("Evidence verification failed: %s", exc)
            return evidence  # Return original if verification fails

    def _load_prompt(self) -> str:
        """Load and cache the prompt template from the ``prompts/`` directory."""
        if self._prompt_template is not None:
            return self._prompt_template

        path = _PROMPTS_DIR / self.prompt_file
        if not path.exists():
            raise FileNotFoundError(
                f"Prompt template not found: {path}"
            )
        self._prompt_template = path.read_text(encoding="utf-8")
        return self._prompt_template

    def _parse_response(self, content: str) -> dict:
        """Parse the LLM JSON response, extracting score/feedback/evidence.

        Handles markdown code fences (````` ```json ... ``` `````) and
        validates the score range (0-4).
        """
        text = content.strip()

        # Strip markdown code fences
        fence_match = re.search(
            r"```(?:json)?\s*\n?(.*?)\n?\s*```", text, re.DOTALL
        )
        if fence_match:
            text = fence_match.group(1).strip()

        # Try standard JSON parse first, fall back to json_repair for
        # malformed output (common with Chinese/non-English LLM responses)
        try:
            data: dict = json.loads(text)
        except json.JSONDecodeError:
            try:
                from json_repair import repair_json
                repaired = repair_json(text, return_objects=True)
                if isinstance(repaired, dict):
                    data = repaired
                else:
                    raise ValueError("Repaired JSON is not a dict")
            except Exception:
                # If repair also fails, re-raise original error
                data = json.loads(text)  # will raise JSONDecodeError

        # Normalise score
        raw_score = data.get("score")
        if raw_score is not None:
            score = int(raw_score)
            if not 0 <= score <= 4:
                raise ValueError(
                    f"Score {score} is out of the valid range 0-4"
                )
            data["score"] = score

        # Normalise evidence to a list of strings
        evidence = data.get("evidence", [])
        if isinstance(evidence, str):
            data["evidence"] = [evidence]

        return data
