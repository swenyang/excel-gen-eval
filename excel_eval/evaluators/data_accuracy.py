"""Data Accuracy evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.evaluators.context_helpers import (
    build_diff_context, estimate_tokens, should_use_diff_mode, smart_text,
)
from excel_eval.models import DimensionName, PreparedData, Scenario

# Token budget for data context (excluding scan report and prompt)
# Leave room for scan report (~3K), prompt (~2K), and LLM response (~4K)
MAX_DATA_TOKENS = 80_000

# Keep these as aliases for backward compatibility (completeness.py imports them)
_estimate_tokens = estimate_tokens
_smart_text = smart_text


class DataAccuracyEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.DATA_ACCURACY

    @property
    def prompt_file(self) -> str:
        return "data_accuracy.md"

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        # Use diff-only mode for large files with high match rate
        if should_use_diff_mode(data):
            return build_diff_context(data, include_generated_sample=True)

        parts: list[str] = []

        # Code-level scan report (FACTUAL — top priority, always included)
        if data.scan_report_text:
            parts.append(data.scan_report_text)

        # Calculate remaining token budget for source + generated data
        scan_tokens = _estimate_tokens(data.scan_report_text) if data.scan_report_text else 0
        remaining_budget = MAX_DATA_TOKENS - scan_tokens

        # Split budget: ~50% source, ~50% generated
        source_budget = remaining_budget // 2
        generated_budget = remaining_budget - source_budget

        # Source data — full if fits, sampled if not
        source_mode = "none"
        if data.grounding_data:
            text, source_mode = _smart_text(
                data.grounding_data, source_budget, "Source Data (Grounding)"
            )
            parts.append(text)

        # Generated Excel content — full if fits, sampled if not
        generated_parts: list[str] = []
        total_gen_tokens = sum(_estimate_tokens(s.csv_text) for s in data.visible_sheets)

        if total_gen_tokens <= generated_budget:
            # Full data fits — send everything
            generated_parts.append(f"## Generated Excel Content — 全量 ({total_gen_tokens} est. tokens)")
            for sheet in data.visible_sheets:
                generated_parts.append(
                    f"### Sheet: {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols)\n{sheet.csv_text}"
                )
        else:
            # Too large — smart sampling per sheet
            per_sheet_budget = generated_budget // max(len(data.visible_sheets), 1)
            generated_parts.append(f"## Generated Excel Content — 采样 (原始 {total_gen_tokens} tokens)")
            for sheet in data.visible_sheets:
                sheet_tokens = _estimate_tokens(sheet.csv_text)
                if sheet_tokens <= per_sheet_budget:
                    generated_parts.append(
                        f"### Sheet: {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols) [全量]\n{sheet.csv_text}"
                    )
                else:
                    lines = sheet.csv_text.split("\n")
                    head_n = min(30, len(lines) // 3)
                    tail_n = min(15, len(lines) // 4)
                    omitted = len(lines) - head_n - tail_n
                    sample = "\n".join(
                        lines[:head_n]
                        + [f"[... {omitted} rows omitted ...]"]
                        + lines[-tail_n:]
                    )
                    generated_parts.append(
                        f"### Sheet: {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols) [采样]\n{sample}"
                    )

        parts.extend(generated_parts)

        # Add a coverage summary so LLM knows what it's working with
        parts.append(
            f"\n---\n**Coverage note**: Source data = {source_mode}, "
            f"Generated data = {'full' if total_gen_tokens <= generated_budget else 'sampled'}. "
            f"{'Rely on the Code-Level Scan Report for full-data comparison.' if data.scan_report_text else 'No code-level scan available — base assessment on visible data only.'}"
        )

        return "\n\n".join(parts)
