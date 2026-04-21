"""Shared context helpers for evaluators.

Provides diff-only mode for cases where input and output are nearly
identical (e.g., audit/debug/template-completion tasks). When the scanner
shows high match rate and data is large, we send only the scan report +
diff details instead of full CSV data — dramatically reducing token usage.
"""

from __future__ import annotations

from excel_eval.models import PreparedData

CHARS_PER_TOKEN = 4  # rough estimate for mixed text/CSV
DIFF_MODE_MATCH_THRESHOLD = 0.70  # match rate above this triggers diff mode
DIFF_MODE_TOKEN_THRESHOLD = 50_000  # only use diff mode if full data exceeds this


def estimate_tokens(text: str) -> int:
    """Rough token estimate for mixed text/CSV content."""
    return len(text) // CHARS_PER_TOKEN


DIFF_MODE_MAX_EXPECTED_FILLS = 100  # don't use diff mode if too many cells were filled


def should_use_diff_mode(data: PreparedData) -> bool:
    """Check if diff-only mode should be used instead of full CSV.

    Diff mode is used when input and output files are structurally similar
    AND the changes are small (e.g., audit/debug tasks that fix errors).

    NOT used for completion tasks (financial modeling, template filling)
    where many empty cells get populated — the LLM needs to see the filled
    values to evaluate accuracy and completeness.

    Conditions (all must be true):
    - Scanner comparison exists (input and output share sheet structure)
    - Grounding data (input) and generated data are similar in size
    - Full data exceeds token threshold
    - Few expected fills (< 100 cells transitioned from empty → value)
    """
    if not data.scan_report_text:
        return False
    if not data.grounding_data:
        return False

    grounding_tokens = estimate_tokens(data.grounding_data)
    generated_tokens = sum(estimate_tokens(s.csv_text) for s in data.visible_sheets)
    total = grounding_tokens + generated_tokens

    # Not worth optimizing if data is small
    if total < DIFF_MODE_TOKEN_THRESHOLD:
        return False

    # If grounding and generated are similar size, it's likely a modify task
    if grounding_tokens == 0 or generated_tokens == 0:
        return False
    ratio = max(grounding_tokens, generated_tokens) / min(grounding_tokens, generated_tokens)
    if ratio >= 3.0:
        return False

    # Don't use diff mode if many cells were filled (completion task, not audit).
    # The LLM needs to see the filled values to evaluate accuracy.
    import re
    fills_match = re.search(
        r"Expected fills.*?(\d+)\s*cells", data.scan_report_text
    )
    if fills_match and int(fills_match.group(1)) > DIFF_MODE_MAX_EXPECTED_FILLS:
        return False

    return True


def build_diff_context(data: PreparedData, include_generated_sample: bool = False) -> str:
    """Build a compact context using only scan report + diff details.

    Used when input and output are nearly identical (audit/debug/template tasks).
    Instead of sending full CSV data, we send:
    1. The scan report (sheet profiles, match rate, actual diffs, expected fills)
    2. Optionally, a small sample of generated data for structural reference

    Args:
        data: Prepared data with scan_report_text populated
        include_generated_sample: If True, include first 10 rows of each sheet
    """
    parts: list[str] = []

    # Scan report is the primary data source in diff mode
    parts.append(data.scan_report_text)

    parts.append(
        "\n---\n"
        "**Context mode: diff-only.** Input and output data are largely identical "
        "(high match rate detected). Only the scan report with differences and "
        "sheet profiles are shown above. Full CSV data is omitted to reduce context size. "
        "Base your assessment on the scan report findings, diff details, and expected fills."
    )

    # Optional: small sample for structural reference (headers + a few rows)
    if include_generated_sample and data.visible_sheets:
        parts.append("## Generated Excel — Structure Sample (first 10 rows per sheet)")
        for sheet in data.visible_sheets:
            lines = sheet.csv_text.split("\n")
            sample = "\n".join(lines[:min(11, len(lines))])  # header + 10 rows
            parts.append(
                f"### {sheet.name} ({sheet.row_count} rows × {sheet.col_count} cols)\n{sample}"
            )

    return "\n\n".join(parts)


def smart_text(text: str, budget_tokens: int, label: str) -> tuple[str, str]:
    """Return (formatted_text, mode) where mode is 'full' or 'sampled'.

    If text fits within budget, returns full text. Otherwise, samples head + tail.
    """
    est = estimate_tokens(text)
    if est <= budget_tokens:
        return f"## {label} — 全量 ({est} est. tokens)\n{text}", "full"

    lines = text.split("\n")
    head_n = min(60, len(lines) // 3)
    tail_n = min(30, len(lines) // 4)
    omitted = len(lines) - head_n - tail_n
    sample = "\n".join(
        lines[:head_n]
        + ["", f"[... {omitted} lines omitted for context budget ...]", ""]
        + lines[-tail_n:]
    )
    return f"## {label} — 采样 (原始 {est} tokens, 已截取首尾)\n{sample}", "sampled"
