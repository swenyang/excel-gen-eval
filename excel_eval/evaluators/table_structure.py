"""Table Structure evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.evaluators.context_helpers import (
    build_diff_context, should_use_diff_mode,
)
from excel_eval.models import DimensionName, PreparedData, Scenario


class TableStructureEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.TABLE_STRUCTURE

    @property
    def prompt_file(self) -> str:
        return "table_structure.md"

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        # Use diff-only mode for large files with high match rate
        if should_use_diff_mode(data):
            ctx = build_diff_context(data, include_generated_sample=True)
            # Add formatting metadata relevant to table structure
            fmt = data.formatting
            fmt_parts: list[str] = ["\n## Table Structure Metadata"]
            if fmt.merged_cell_ranges:
                fmt_parts.append(f"### Merged Cells ({len(fmt.merged_cell_ranges)})")
                for mc in fmt.merged_cell_ranges[:30]:
                    fmt_parts.append(f"- {mc}")
            else:
                fmt_parts.append("### Merged Cells\nNone")
            if fmt.frozen_panes:
                fmt_parts.append("### Frozen Panes")
                for sheet_name, freeze_point in fmt.frozen_panes.items():
                    fmt_parts.append(f"- {sheet_name}: frozen at {freeze_point}")
            else:
                fmt_parts.append("### Frozen Panes\nNone")
            return ctx + "\n\n".join(fmt_parts)

        parts: list[str] = []

        # CSV content
        parts.append("## Sheet Content")
        for sheet in data.visible_sheets:
            parts.append(f"### Sheet: {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols)")
            parts.append(sheet.csv_text)

        # Formatting metadata relevant to table structure
        fmt = data.formatting
        parts.append("## Table Structure Metadata")

        if fmt.merged_cell_ranges:
            parts.append(f"### Merged Cells ({len(fmt.merged_cell_ranges)})")
            for mc in fmt.merged_cell_ranges[:30]:
                parts.append(f"- {mc}")
        else:
            parts.append("### Merged Cells\nNone")

        if fmt.frozen_panes:
            parts.append("### Frozen Panes")
            for sheet_name, freeze_point in fmt.frozen_panes.items():
                parts.append(f"- {sheet_name}: frozen at {freeze_point}")
        else:
            parts.append("### Frozen Panes\nNone")

        return "\n\n".join(parts)
