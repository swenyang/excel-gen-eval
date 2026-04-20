"""Relevance evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.evaluators.context_helpers import (
    build_diff_context, should_use_diff_mode,
)
from excel_eval.models import DimensionName, PreparedData, Scenario


class RelevanceEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.RELEVANCE

    @property
    def prompt_file(self) -> str:
        return "relevance.md"

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        # Use diff-only mode for large files with high match rate
        if should_use_diff_mode(data):
            return build_diff_context(data, include_generated_sample=True)

        parts: list[str] = []

        # Generated Excel content
        parts.append("## Generated Excel Content")
        for sheet in data.visible_sheets:
            parts.append(f"### Sheet: {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols)")
            parts.append(sheet.csv_text)

        return "\n\n".join(parts)
