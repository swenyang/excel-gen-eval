"""Completeness evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.evaluators.data_accuracy import _estimate_tokens, _smart_text
from excel_eval.models import DimensionName, PreparedData, Scenario

MAX_DATA_TOKENS = 80_000


class CompletenessEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.COMPLETENESS

    @property
    def prompt_file(self) -> str:
        return "completeness.md"

    def needs_screenshots(self) -> bool:
        return True

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        parts: list[str] = []

        # Code-level scan report
        if data.scan_report_text:
            parts.append(data.scan_report_text)

        scan_tokens = _estimate_tokens(data.scan_report_text) if data.scan_report_text else 0
        remaining = MAX_DATA_TOKENS - scan_tokens

        # Source data — full if fits
        if data.grounding_data:
            text, _ = _smart_text(data.grounding_data, remaining // 2, "Source Data (Grounding)")
            parts.append(text)
            remaining -= _estimate_tokens(text)

        # Sheet list
        sheet_names = [s.name for s in data.visible_sheets]
        parts.append(f"## Sheet List\n{', '.join(sheet_names)}")

        # Generated Excel content — full if fits
        total_gen_tokens = sum(_estimate_tokens(s.csv_text) for s in data.visible_sheets)
        if total_gen_tokens <= remaining:
            parts.append("## Generated Excel Content — 全量")
            for sheet in data.visible_sheets:
                parts.append(f"### Sheet: {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols)\n{sheet.csv_text}")
        else:
            parts.append("## Generated Excel Content — 采样")
            per_sheet = remaining // max(len(data.visible_sheets), 1)
            for sheet in data.visible_sheets:
                if _estimate_tokens(sheet.csv_text) <= per_sheet:
                    parts.append(f"### {sheet.name} ({sheet.row_count}×{sheet.col_count}) [全量]\n{sheet.csv_text}")
                else:
                    lines = sheet.csv_text.split("\n")
                    head_n, tail_n = min(25, len(lines)//3), min(10, len(lines)//4)
                    sample = "\n".join(lines[:head_n] + [f"[... {len(lines)-head_n-tail_n} rows omitted ...]"] + lines[-tail_n:])
                    parts.append(f"### {sheet.name} ({sheet.row_count}×{sheet.col_count}) [采样]\n{sample}")

        # Chart summary
        if data.charts:
            chart_summary = "\n".join(
                f"- {c.chart_type} chart on sheet '{c.sheet}': {c.title or '(untitled)'}"
                for c in data.charts
            )
            parts.append(f"## Charts Present\n{chart_summary}")

        return "\n\n".join(parts)
