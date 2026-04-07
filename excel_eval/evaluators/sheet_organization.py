"""Sheet Organization evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.models import DimensionName, PreparedData, Scenario


class SheetOrganizationEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.SHEET_ORGANIZATION

    @property
    def prompt_file(self) -> str:
        return "sheet_organization.md"

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        parts: list[str] = []

        # Sheet names and metadata
        parts.append("## Sheet Structure")
        for i, sheet in enumerate(data.sheets, 1):
            parts.append(f"{i}. **{sheet.name}** — {sheet.row_count} rows × {sheet.col_count} cols")

        # Cross-sheet references
        if data.cross_sheet_refs:
            parts.append(f"\n## Cross-Sheet References ({len(data.cross_sheet_refs)})")
            for ref in data.cross_sheet_refs[:50]:
                parts.append(f"- {ref}")
        else:
            parts.append("\n## Cross-Sheet References\nNone found.")

        return "\n\n".join(parts)
