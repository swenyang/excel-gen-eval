"""Formula & Logic evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.models import DimensionName, PreparedData, Scenario


class FormulaLogicEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.FORMULA_LOGIC

    @property
    def prompt_file(self) -> str:
        return "formula_logic.md"

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        parts: list[str] = []

        # Formula metadata
        if data.formulas:
            parts.append(f"## Formula Inventory ({len(data.formulas)} formulas)")
            error_count = sum(1 for f in data.formulas if f.has_error)
            parts.append(f"- Total formulas: {len(data.formulas)}")
            parts.append(f"- Formulas with errors: {error_count}")
            parts.append("")

            parts.append("### Formula Details")
            for f in data.formulas[:200]:  # Cap at 200 formulas
                status = " ⚠ ERROR" if f.has_error else ""
                parts.append(
                    f"- `{f.sheet}!{f.cell}`: `{f.formula}` → {f.computed_value}{status}"
                )
            if len(data.formulas) > 200:
                parts.append(f"*(... and {len(data.formulas) - 200} more)*")
        else:
            parts.append("## Formula Inventory\n**No formulas found in this workbook.**")

        # CSV content for context
        parts.append("\n## Sheet Content (for reference)")
        for sheet in data.visible_sheets:
            lines = sheet.csv_text.splitlines()
            preview = "\n".join(lines[:30])  # First 30 rows for context
            parts.append(f"### {sheet.name}\n{preview}")

        return "\n\n".join(parts)
