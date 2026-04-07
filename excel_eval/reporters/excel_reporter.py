"""Excel reporter — writes evaluation results as an Excel workbook."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from ..models import DimensionName, EvalResult, EvalStatus


def generate_excel_report(
    results: list[EvalResult],
    output_path: str | Path,
) -> Path:
    """Write evaluation results as an Excel file with multiple sheets."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = xlsxwriter.Workbook(str(output_path))

    # Formats
    header_fmt = wb.add_format({
        "bold": True, "bg_color": "#4472C4", "font_color": "white",
        "border": 1, "text_wrap": True, "valign": "vcenter",
    })
    score_fmts = {
        0: wb.add_format({"bg_color": "#FF6B6B", "align": "center", "border": 1}),
        1: wb.add_format({"bg_color": "#FFA07A", "align": "center", "border": 1}),
        2: wb.add_format({"bg_color": "#FFD700", "align": "center", "border": 1}),
        3: wb.add_format({"bg_color": "#90EE90", "align": "center", "border": 1}),
        4: wb.add_format({"bg_color": "#3CB371", "align": "center", "border": 1}),
    }
    na_fmt = wb.add_format({"bg_color": "#D3D3D3", "align": "center", "border": 1})
    text_fmt = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
    num_fmt = wb.add_format({"align": "center", "border": 1, "num_format": "0.00"})

    # Sheet 1: Score Summary
    _write_summary_sheet(wb, results, header_fmt, score_fmts, na_fmt, num_fmt)

    # Sheet 2: Dimension Details
    _write_details_sheet(wb, results, header_fmt, score_fmts, na_fmt, text_fmt)

    # Sheet 3: Cost Summary
    _write_cost_sheet(wb, results, header_fmt, text_fmt, num_fmt)

    wb.close()
    return output_path


def _write_summary_sheet(wb, results, header_fmt, score_fmts, na_fmt, num_fmt):
    ws = wb.add_worksheet("Score Summary")

    dims = [d.value for d in DimensionName]
    headers = ["Case ID", "Scenario"] + dims + [
        "Data & Content Avg", "Structure & Usability Avg", "Overall Weighted Avg"
    ]

    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    for row, result in enumerate(results, 1):
        ws.write(row, 0, result.case_id)
        ws.write(row, 1, result.scenario.detected.value)

        for col, dim in enumerate(dims, 2):
            dr = result.dimensions.get(dim)
            if dr and dr.score is not None:
                fmt = score_fmts.get(dr.score, na_fmt)
                ws.write(row, col, dr.score, fmt)
            else:
                ws.write(row, col, "N/A", na_fmt)

        avg_col = len(dims) + 2
        ws.write(row, avg_col, result.summary.data_content_avg or 0, num_fmt)
        ws.write(row, avg_col + 1, result.summary.structure_usability_avg or 0, num_fmt)
        ws.write(row, avg_col + 2, result.summary.overall_weighted_avg or 0, num_fmt)

    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 22)
    ws.set_column(2, len(headers) - 1, 16)


def _write_details_sheet(wb, results, header_fmt, score_fmts, na_fmt, text_fmt):
    ws = wb.add_worksheet("Dimension Details")

    headers = ["Case ID", "Dimension", "Score", "Status", "Feedback", "Evidence"]
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    row = 1
    for result in results:
        for dim_name, dr in result.dimensions.items():
            ws.write(row, 0, result.case_id)
            ws.write(row, 1, dim_name)

            if dr.score is not None:
                fmt = score_fmts.get(dr.score, na_fmt)
                ws.write(row, 2, dr.score, fmt)
            else:
                ws.write(row, 2, "N/A", na_fmt)

            ws.write(row, 3, dr.status.value)
            ws.write(row, 4, dr.feedback, text_fmt)
            ws.write(row, 5, "\n".join(dr.evidence), text_fmt)
            row += 1

    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 25)
    ws.set_column(2, 3, 10)
    ws.set_column(4, 4, 60)
    ws.set_column(5, 5, 60)


def _write_cost_sheet(wb, results, header_fmt, text_fmt, num_fmt):
    ws = wb.add_worksheet("Cost Summary")

    headers = ["Case ID", "Total Input Tokens", "Total Output Tokens",
               "Total Latency (ms)", "Est. Cost ($)"]
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)

    for row, result in enumerate(results, 1):
        ws.write(row, 0, result.case_id)
        ws.write(row, 1, result.cost.total_input_tokens, num_fmt)
        ws.write(row, 2, result.cost.total_output_tokens, num_fmt)
        ws.write(row, 3, result.cost.total_latency_ms, num_fmt)
        ws.write(row, 4, result.cost.total_cost_estimate, num_fmt)

    ws.set_column(0, 0, 30)
    ws.set_column(1, 4, 20)
