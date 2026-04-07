"""Professional Formatting evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.models import DimensionName, PreparedData, Scenario


class ProfessionalFormattingEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.PROFESSIONAL_FORMATTING

    @property
    def prompt_file(self) -> str:
        return "professional_formatting.md"

    def needs_screenshots(self) -> bool:
        return True

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        parts: list[str] = []

        fmt = data.formatting

        # Fonts
        parts.append("## Formatting Metadata")
        parts.append(f"### Fonts Used\n{', '.join(fmt.fonts_used) if fmt.fonts_used else 'Default only'}")

        # Color palette
        if fmt.color_palette:
            parts.append(f"### Color Palette ({len(fmt.color_palette)} colors)\n{', '.join(fmt.color_palette)}")

        # Conditional formatting
        if fmt.has_conditional_formatting:
            parts.append(f"### Conditional Formatting ({len(fmt.conditional_format_rules)} rules)")
            for rule in fmt.conditional_format_rules[:20]:
                parts.append(f"- {rule}")
        else:
            parts.append("### Conditional Formatting\nNone applied")

        # Merged cells
        if fmt.merged_cell_ranges:
            parts.append(f"### Merged Cells ({len(fmt.merged_cell_ranges)})")
            for mc in fmt.merged_cell_ranges[:20]:
                parts.append(f"- {mc}")

        # Frozen panes
        if fmt.frozen_panes:
            parts.append("### Frozen Panes")
            for sheet_name, fp in fmt.frozen_panes.items():
                parts.append(f"- {sheet_name}: {fp}")

        # Sheet content for context
        parts.append("\n## Sheet Content (for context)")
        for sheet in data.sheets:
            lines = sheet.csv_text.splitlines()
            preview = "\n".join(lines[:15])
            parts.append(f"### {sheet.name}\n{preview}")

        return "\n\n".join(parts)
