"""Chart Appropriateness evaluator."""

from __future__ import annotations

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.models import DimensionName, PreparedData, Scenario


class ChartAppropriatenessEvaluator(BaseEvaluator):

    @property
    def dimension(self) -> DimensionName:
        return DimensionName.CHART_APPROPRIATENESS

    @property
    def prompt_file(self) -> str:
        return "chart_appropriateness.md"

    def needs_screenshots(self) -> bool:
        return True

    def build_context(self, data: PreparedData, scenario: Scenario) -> str:
        parts: list[str] = []

        # Chart metadata
        if data.charts:
            parts.append(f"## Charts ({len(data.charts)})")
            for i, chart in enumerate(data.charts, 1):
                parts.append(f"### Chart {i}")
                parts.append(f"- **Sheet**: {chart.sheet}")
                parts.append(f"- **Type**: {chart.chart_type}")
                parts.append(f"- **Title**: {chart.title or '(none)'}")
                parts.append(f"- **Data range**: {chart.data_range or '(unknown)'}")
                parts.append(f"- **Has legend**: {chart.has_legend}")
                parts.append(f"- **Has axis labels**: {chart.has_axis_labels}")
        else:
            parts.append("## Charts\n**No charts found in this workbook.**")

        # CSV content for reference
        parts.append("\n## Sheet Content (for data reference)")
        for sheet in data.visible_sheets:
            lines = sheet.csv_text.splitlines()
            preview = "\n".join(lines[:20])
            parts.append(f"### {sheet.name}\n{preview}")

        return "\n\n".join(parts)
