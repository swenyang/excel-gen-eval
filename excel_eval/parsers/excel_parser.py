"""Excel file parser — extracts CSV, formulas, charts, and formatting metadata."""

from __future__ import annotations

import io
import re
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, ScatterChart
from openpyxl.utils import get_column_letter

from ..models import (
    ChartInfo,
    FormatInfo,
    FormulaInfo,
    PreparedData,
    SheetData,
)

# Row truncation thresholds
MAX_ROWS_FULL = 500
HEAD_ROWS = 100
TAIL_ROWS = 50


def parse_excel(
    excel_path: str | Path,
    grounding_data: str = "",
    user_prompt: str = "",
) -> PreparedData:
    """Parse an Excel file and extract all structured data for evaluation.

    Returns a PreparedData object with sheets, formulas, charts,
    formatting metadata, and cross-sheet references.
    """
    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    wb = openpyxl.load_workbook(str(excel_path), data_only=False)
    wb_values = openpyxl.load_workbook(str(excel_path), data_only=True)

    sheets = _extract_sheets(excel_path)
    formulas = _extract_formulas(wb, wb_values)
    charts = _extract_charts(wb)
    formatting = _extract_formatting(wb)
    cross_refs = _extract_cross_sheet_refs(wb)

    wb.close()
    wb_values.close()

    return PreparedData(
        sheets=sheets,
        formulas=formulas,
        charts=charts,
        formatting=formatting,
        cross_sheet_refs=cross_refs,
        screenshots={},  # Populated separately by screenshot module
        grounding_data=grounding_data,
        user_prompt=user_prompt,
    )


def _extract_sheets(excel_path: Path) -> list[SheetData]:
    """Export each sheet to CSV text using Excel's display format.

    Reads cell number_format via openpyxl to produce values matching
    what the user sees in Excel, avoiding CSV export artifacts like
    datetime timestamps or floating-point noise.
    """
    wb_display = openpyxl.load_workbook(str(excel_path), data_only=True)
    xls = pd.ExcelFile(excel_path)
    sheets: list[SheetData] = []

    for sheet_name in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        row_count, col_count = df_raw.shape

        # Build display-formatted DataFrame from openpyxl
        ws = wb_display[sheet_name]
        rows_data = []
        headers = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=False)):
            if row_idx == 0:
                headers = [_display_value(cell) for cell in row]
                continue
            rows_data.append([_display_value(cell) for cell in row])

        df = pd.DataFrame(rows_data, columns=headers) if headers else df_raw

        truncated = False
        if row_count > MAX_ROWS_FULL:
            head = df.head(HEAD_ROWS)
            tail = df.tail(TAIL_ROWS)
            truncated_count = row_count - HEAD_ROWS - TAIL_ROWS
            marker = pd.DataFrame(
                {col: [f"[... {truncated_count} rows truncated ...]"] for col in df.columns},
                index=[0],
            )
            df = pd.concat([head, marker, tail], ignore_index=True)
            truncated = True

        csv_text = df.to_csv(index=False)
        sheets.append(SheetData(
            name=sheet_name,
            csv_text=csv_text,
            row_count=row_count,
            col_count=col_count,
            truncated=truncated,
        ))

    wb_display.close()
    return sheets


def _display_value(cell) -> str:
    """Format a cell value based on its Excel number_format.

    Produces the value as the user would see it in Excel, not the raw
    underlying value.
    """
    val = cell.value
    if val is None:
        return ""

    fmt = cell.number_format or "General"

    # Date/time formats — detect by common Excel format tokens
    if isinstance(val, datetime):
        if any(tok in fmt.lower() for tok in ["h:", "hh:", "ss", "am/pm"]):
            return val.strftime("%Y-%m-%d %H:%M:%S")
        else:
            return val.strftime("%Y-%m-%d")

    # Numeric — format to match Excel display
    if isinstance(val, (int, float)):
        use_thousands = "," in fmt or "#,#" in fmt

        # Determine decimal places
        if "." in fmt:
            decimal_part = fmt.split(".")[-1]
            decimals = len(decimal_part.replace("0", "X").replace("#", "X").split(";")[0].rstrip(")%"))
            decimals = min(decimals, 10)
        elif fmt == "General":
            decimals = 6 if isinstance(val, float) else 0
        else:
            decimals = 0

        rounded = round(float(val), decimals)

        # Format with or without thousand separators
        if decimals == 0:
            int_val = int(rounded)
            if use_thousands:
                return f"{int_val:,}"
            return str(int_val)
        else:
            if use_thousands:
                return f"{rounded:,.{decimals}f}"
            return f"{rounded:.{decimals}f}"

    return str(val)


def _extract_formulas(
    wb: openpyxl.Workbook,
    wb_values: openpyxl.Workbook,
) -> list[FormulaInfo]:
    """Extract all formulas with their computed values."""
    formulas: list[FormulaInfo] = []

    for ws in wb.worksheets:
        ws_values = wb_values[ws.title]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    coord = cell.coordinate
                    computed = ws_values[coord].value
                    has_error = isinstance(computed, str) and computed.startswith("#")

                    formulas.append(FormulaInfo(
                        cell=coord,
                        sheet=ws.title,
                        formula=cell.value,
                        computed_value=computed,
                        has_error=has_error,
                    ))

    return formulas


def _extract_charts(wb: openpyxl.Workbook) -> list[ChartInfo]:
    """Extract chart metadata from all sheets."""
    charts: list[ChartInfo] = []
    chart_type_map = {
        BarChart: "bar",
        LineChart: "line",
        PieChart: "pie",
        AreaChart: "area",
        ScatterChart: "scatter",
    }

    for ws in wb.worksheets:
        for chart in ws._charts:
            chart_type = "unknown"
            for cls, name in chart_type_map.items():
                if isinstance(chart, cls):
                    chart_type = name
                    break

            title = None
            if chart.title:
                title = str(chart.title)

            data_range = None
            if chart.series:
                try:
                    data_range = str(chart.series[0].val.numRef.f)
                except (AttributeError, IndexError):
                    pass

            has_legend = chart.legend is not None
            has_axis_labels = False
            if hasattr(chart, "x_axis") and chart.x_axis.title:
                has_axis_labels = True
            elif hasattr(chart, "y_axis") and chart.y_axis.title:
                has_axis_labels = True

            charts.append(ChartInfo(
                sheet=ws.title,
                chart_type=chart_type,
                title=title,
                data_range=data_range,
                has_legend=has_legend,
                has_axis_labels=has_axis_labels,
            ))

    return charts


def _extract_formatting(wb: openpyxl.Workbook) -> FormatInfo:
    """Extract workbook-level formatting metadata."""
    fonts_used: set[str] = set()
    colors_used: set[str] = set()
    merged_ranges: list[str] = []
    conditional_rules: list[str] = []
    frozen_panes: dict[str, str] = {}
    has_cf = False

    for ws in wb.worksheets:
        # Fonts and colors
        for row in ws.iter_rows(max_row=min(ws.max_row or 0, 100)):
            for cell in row:
                if cell.font and cell.font.name:
                    fonts_used.add(cell.font.name)
                if cell.font and cell.font.color and cell.font.color.rgb:
                    color = str(cell.font.color.rgb)
                    if color != "00000000":
                        colors_used.add(color)
                if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                    color = str(cell.fill.start_color.rgb)
                    if color != "00000000":
                        colors_used.add(color)

        # Merged cells
        for merged_range in ws.merged_cells.ranges:
            merged_ranges.append(f"{ws.title}!{merged_range}")

        # Conditional formatting
        for cf_rule in ws.conditional_formatting:
            has_cf = True
            conditional_rules.append(f"{ws.title}: {cf_rule}")

        # Frozen panes
        if ws.freeze_panes:
            frozen_panes[ws.title] = str(ws.freeze_panes)

    return FormatInfo(
        fonts_used=sorted(fonts_used),
        color_palette=sorted(colors_used)[:20],  # Cap at 20 colors
        has_conditional_formatting=has_cf,
        conditional_format_rules=conditional_rules[:50],  # Cap at 50 rules
        merged_cell_ranges=merged_ranges,
        frozen_panes=frozen_panes,
    )


def _extract_cross_sheet_refs(wb: openpyxl.Workbook) -> list[str]:
    """Find formulas that reference other sheets.

    Detects:
    - Direct references: Sheet2!A1, 'Sheet Name'!A1
    - Defined names that span sheets
    - INDIRECT() references (flagged as potential cross-sheet)
    """
    cross_refs: list[str] = []
    sheet_names = set(ws.title for ws in wb.worksheets)

    # Check defined names for cross-sheet scope
    try:
        defined = wb.defined_names
        # openpyxl versions differ: try .definedName, then iterate directly
        names_iter = getattr(defined, 'definedName', None) or defined.values()
        for name in names_iter:
            try:
                destinations = list(name.destinations)
                if len(destinations) > 1:
                    sheets_involved = [d[0] for d in destinations]
                    cross_refs.append(
                        f"Defined name '{name.name}' spans sheets: {', '.join(sheets_involved)}"
                    )
                elif destinations and name.attr_text and "!" in name.attr_text:
                    cross_refs.append(
                        f"Defined name '{name.name}' → {name.attr_text}"
                    )
            except Exception:
                pass
    except Exception:
        pass

    # Check formulas for direct sheet references
    seen = set()  # Deduplicate (only track unique source→target pairs)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue
                formula = cell.value

                for other_sheet in sheet_names:
                    if other_sheet == ws.title:
                        continue
                    # Match: SheetName!Cell, 'Sheet Name'!Cell
                    patterns = [f"{other_sheet}!", f"'{other_sheet}'!"]
                    for pattern in patterns:
                        if pattern in formula:
                            pair = (ws.title, other_sheet)
                            if pair not in seen:
                                seen.add(pair)
                                cross_refs.append(
                                    f"{ws.title} → {other_sheet} "
                                    f"(e.g., {cell.coordinate}: {formula[:80]})"
                                )
                            break

                # Flag INDIRECT as potential cross-sheet
                if "INDIRECT(" in formula.upper() and (ws.title, "_INDIRECT") not in seen:
                    seen.add((ws.title, "_INDIRECT"))
                    cross_refs.append(
                        f"{ws.title}!{cell.coordinate} uses INDIRECT() "
                        f"(potential cross-sheet: {formula[:80]})"
                    )

    return cross_refs
