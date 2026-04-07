"""Relevance evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.models import DimensionName, PreparedData, Scenario


class RelevanceEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.RELEVANCE

    @property
    def prompt_file(self) -> str:
        return "relevance.md"

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        parts: list[str] = []

        # Generated Excel content
        parts.append("## Generated Excel Content")
        for sheet in data.sheets:
            parts.append(f"### Sheet: {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols)")
            parts.append(sheet.csv_text)

        return "\n\n".join(parts)
