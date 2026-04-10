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
        for i, sheet in enumerate(data.visible_sheets, 1):
            parts.append(f"{i}. **{sheet.name}** — {sheet.row_count} rows × {sheet.col_count} cols")

            # Detect sections within each sheet (blank rows = section separators)
            sections = _detect_sections(sheet.csv_text)
            if sections and len(sections) > 1:
                parts.append(f"   Sections detected ({len(sections)}):")
                for sec in sections:
                    parts.append(f"   - Rows {sec['start']}-{sec['end']}: {sec['label']}")

        # Cross-sheet references
        if data.cross_sheet_refs:
            parts.append(f"\n## Cross-Sheet References ({len(data.cross_sheet_refs)})")
            for ref in data.cross_sheet_refs[:50]:
                parts.append(f"- {ref}")
        else:
            parts.append("\n## Cross-Sheet References\nNone found.")

        return "\n\n".join(parts)


def _detect_sections(csv_text: str) -> list[dict]:
    """Detect logical sections in a sheet by finding blank-row separators."""
    lines = csv_text.strip().split("\n")
    if len(lines) < 3:
        return []

    sections: list[dict] = []
    current_start = 1  # skip header
    current_label = ""

    for i, line in enumerate(lines[1:], start=2):  # 1-indexed, skip header
        cells = [c.strip() for c in line.split(",")]
        is_blank = all(c == "" or c == '""' for c in cells)
        is_header_like = sum(1 for c in cells if c and not c.replace(".", "").replace("-", "").isdigit()) > len(cells) * 0.6

        if is_blank and current_start < i:
            # End of a section
            label = current_label or f"Data block"
            sections.append({"start": current_start, "end": i - 1, "label": label})
            current_start = i + 1
            current_label = ""
        elif is_header_like and i == current_start:
            # This row looks like a section header
            current_label = line.split(",")[0].strip().strip('"')[:50]

    # Last section
    if current_start < len(lines):
        label = current_label or "Data block"
        sections.append({"start": current_start, "end": len(lines), "label": label})

    return sections
