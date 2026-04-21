"""HTML report generator for Excel evaluation results.

Produces a single self-contained HTML file with two views:
  1. Summary Table — one row per case, all scores at a glance.
  2. Case Detail — shown when clicking a case, with metadata, prompt,
     files, score overview, and highlighted dimension feedback.
"""

from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path

from jinja2 import Environment
from markupsafe import Markup

from excel_eval.models import DimensionName, EvalResult, EvalStatus, SCENARIO_WEIGHTS, get_blended_weights

# ── Dimension display names and grouping ──────────────────────────────────

_DIMENSION_LABELS: dict[str, str] = {
    DimensionName.DATA_ACCURACY: "Data Accuracy",
    DimensionName.COMPLETENESS: "Completeness",
    DimensionName.FORMULA_LOGIC: "Formula Logic",
    DimensionName.RELEVANCE: "Relevance",
    DimensionName.SHEET_ORGANIZATION: "Sheet Organization",
    DimensionName.TABLE_STRUCTURE: "Table Structure",
    DimensionName.CHART_APPROPRIATENESS: "Chart Appropriateness",
    DimensionName.PROFESSIONAL_FORMATTING: "Professional Formatting",
}

_DIMENSION_SHORT: dict[str, str] = {
    DimensionName.DATA_ACCURACY: "Accuracy",
    DimensionName.COMPLETENESS: "Complete",
    DimensionName.FORMULA_LOGIC: "Formula",
    DimensionName.RELEVANCE: "Relevance",
    DimensionName.SHEET_ORGANIZATION: "Sheets",
    DimensionName.TABLE_STRUCTURE: "Tables",
    DimensionName.CHART_APPROPRIATENESS: "Charts",
    DimensionName.PROFESSIONAL_FORMATTING: "Format",
}

_DATA_CONTENT_DIMS = [
    DimensionName.DATA_ACCURACY,
    DimensionName.COMPLETENESS,
    DimensionName.FORMULA_LOGIC,
    DimensionName.RELEVANCE,
]

_STRUCTURE_USABILITY_DIMS = [
    DimensionName.SHEET_ORGANIZATION,
    DimensionName.TABLE_STRUCTURE,
    DimensionName.CHART_APPROPRIATENESS,
    DimensionName.PROFESSIONAL_FORMATTING,
]

_ALL_DIMS = _DATA_CONTENT_DIMS + _STRUCTURE_USABILITY_DIMS

# Well-known metadata keys to display (in order)
_METADATA_KEYS = ["sector", "occupation", "task_id", "source", "category", "difficulty"]


# ── Helper functions ──────────────────────────────────────────────────────


def _score_color(score: int | float | None) -> str:
    """Return CSS color class name for a 0-4 score."""
    if score is None:
        return "score-na"
    s = round(score)
    if s <= 1:
        return "score-red"
    if s == 2:
        return "score-orange"
    if s == 3:
        return "score-green"
    return "score-dark-green"


def _score_pct(score: int | None) -> float:
    """Convert a 0-4 score to a percentage for bar width."""
    if score is None:
        return 0.0
    return max(0.0, min(100.0, score / 4.0 * 100.0))


_GOOD_RE = re.compile(
    r"\b(correct|good|proper|well|excellent|appropriate|clear|consistent|comprehensive"
    r"|strong|solid|accurate|effective|nicely|adequately|complete)\b",
    re.IGNORECASE,
)
_BAD_RE = re.compile(
    r"\b(missing|error|incorrect|poor|lacking|absent|hardcoded|hard-coded"
    r"|inconsistent|inadequate|wrong|fail|broken|unclear|excessive|unnecessary)\b",
    re.IGNORECASE,
)


def _classify_sentiment(text: str) -> str:
    """Classify sentiment from +/- prefix tags in evidence text."""
    stripped = text.strip()
    if stripped.startswith("+"):
        return "good"
    if stripped.startswith("-"):
        return "bad"
    return "neutral"


# Patterns to split feedback on numbered markers:
# (1) ... (2) ..., or 1. ... 2. ..., or 1) ... 2) ...
_NUMBERED_INLINE_RE = re.compile(r"\s*\(\d+\)\s*")
_NUMBERED_LINE_RE = re.compile(r"(?:^|\n)\s*\d+[).]\s+")


def _split_numbered(text: str) -> list[str]:
    """Split text on numbered markers, return list of parts (empty if no split)."""
    # Try inline (1) (2) (3) pattern first
    parts = _NUMBERED_INLINE_RE.split(text.strip())
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) > 1:
        return parts

    # Try line-based 1. 2. 3. pattern
    parts = _NUMBERED_LINE_RE.split(text.strip())
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) > 1:
        return parts

    return []


def _highlight_feedback(text: str) -> str:
    """Convert feedback text to HTML.

    - Splits numbered text into a <ul> list with sentiment coloring
    - Uses +/- prefix for sentiment; falls back to keyword matching
    - Single paragraphs rendered as-is
    """
    if not text:
        return ""

    from markupsafe import escape

    parts = _split_numbered(text)
    if parts:
        items = []
        for part in parts:
            stripped = part.strip()
            if stripped.startswith(("+", "-")):
                sentiment = "good" if stripped.startswith("+") else "bad"
                display = stripped[1:].strip()
            else:
                # Keyword-based fallback for items without +/- prefix
                good_count = len(_GOOD_RE.findall(stripped))
                bad_count = len(_BAD_RE.findall(stripped))
                if bad_count > good_count:
                    sentiment = "bad"
                elif good_count > bad_count:
                    sentiment = "good"
                else:
                    sentiment = "neutral"
                display = stripped
            escaped = str(escape(display))
            css = "evidence-good" if sentiment == "good" else ("evidence-bad" if sentiment == "bad" else "")
            cls = f' class="{css}"' if css else ""
            items.append(f"<li{cls}>{escaped}</li>")
        return "<ul class='feedback-list'>" + "".join(items) + "</ul>"
    else:
        escaped = str(escape(text))
        return f"<p class='feedback'>{escaped}</p>"


def _highlight_evidence(text: str) -> str:
    """Render evidence item with color based on +/- prefix and UNCONFIRMED status."""
    if not text:
        return ""
    from markupsafe import escape

    # UNCONFIRMED items get special styling
    if "UNCONFIRMED" in text.upper():
        display = text.strip()
        if display.startswith(("+", "-")):
            display = display[1:].strip()
        escaped = str(escape(display))
        return f'<span class="evidence-unconfirmed">{escaped}</span>'

    sentiment = _classify_sentiment(text)
    # Strip the leading +/- for display
    display = text.strip()
    if display.startswith(("+", "-")):
        display = display[1:].strip()
    escaped = str(escape(display))
    css = "evidence-good" if sentiment == "good" else ("evidence-bad" if sentiment == "bad" else "")
    if css:
        return f'<span class="{css}">{escaped}</span>'
    return escaped


# ── Inline Jinja2 template ───────────────────────────────────────────────

_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel Gen Eval Report</title>
<style>
  :root {
    --color-fg: #1f2328;
    --color-fg-muted: #656d76;
    --color-bg: #ffffff;
    --color-bg-subtle: #f6f8fa;
    --color-border: #d1d9e0;
    --color-border-muted: #d8dee4;
    --color-accent: #0969da;
    --color-red: #cf222e;
    --color-orange: #bf8700;
    --color-green: #1a7f37;
    --color-dark-green: #116329;
    --radius: 6px;
  }
  *, *::before, *::after { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans", Helvetica, Arial, sans-serif;
    font-size: 14px;
    line-height: 1.5;
    color: var(--color-fg);
    background: var(--color-bg-subtle);
    margin: 0;
    padding: 24px;
  }
  .container { max-width: 1200px; margin: 0 auto; }

  /* View toggling */
  .view-hidden { display: none !important; }

  /* Header */
  .header {
    background: var(--color-bg);
    border: 1px solid var(--color-border);
    border-radius: var(--radius);
    padding: 24px 32px;
    margin-bottom: 24px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
  }
  .header h1 { margin: 0; font-size: 24px; font-weight: 600; }
  .header-meta { color: var(--color-fg-muted); font-size: 13px; }
  .lang-btn {
    background: var(--color-bg-subtle);
    border: 1px solid var(--color-border);
    border-radius: 20px;
    padding: 4px 14px;
    font-size: 12px;
    cursor: pointer;
    font-weight: 500;
  }
  .lang-btn:hover { background: #e8e8e8; }

  /* Back link */
  .back-link {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    color: var(--color-accent);
    text-decoration: none;
    font-size: 14px;
    font-weight: 500;
    margin-bottom: 16px;
    cursor: pointer;
    padding: 4px 0;
  }
  .back-link:hover { text-decoration: underline; }

  /* ── Summary Table ─────────────────────────────────────── */
  .summary-wrapper {
    background: var(--color-bg);
    border: 1px solid var(--color-border);
    border-radius: var(--radius);
    margin-bottom: 24px;
    overflow-x: auto;
  }
  .summary-wrapper h2 {
    margin: 0;
    padding: 16px 20px;
    font-size: 18px;
    font-weight: 600;
    border-bottom: 1px solid var(--color-border-muted);
  }
  .summary-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }
  .summary-table thead th {
    position: sticky;
    top: 0;
    z-index: 2;
    background: var(--color-bg-subtle);
    font-weight: 600;
    white-space: nowrap;
    padding: 10px 12px;
    text-align: center;
    border-bottom: 2px solid var(--color-border);
  }
  .summary-table thead th:first-child,
  .summary-table thead th:nth-child(2) { text-align: left; }
  .summary-table .block-header {
    text-align: center;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: var(--color-fg-muted);
    border-bottom: 1px solid var(--color-border-muted);
  }
  .summary-table tbody tr:nth-child(even) { background: var(--color-bg-subtle); }
  .summary-table tbody tr:hover { background: #eef6ff; }
  .summary-table td {
    padding: 8px 12px;
    text-align: center;
    vertical-align: middle;
    border-bottom: 1px solid var(--color-border-muted);
  }
  .summary-table td:first-child,
  .summary-table td:nth-child(2) { text-align: left; }
  .summary-table td:last-child { font-weight: 700; }
  .summary-table tr:last-child td { border-bottom: none; }
  .summary-table .case-link {
    color: var(--color-accent);
    text-decoration: none;
    font-weight: 600;
    cursor: pointer;
  }
  .summary-table .case-link:hover { text-decoration: underline; }
  .summary-table .score-cell {
    display: inline-block;
    min-width: 28px;
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 600;
    color: #fff;
    text-align: center;
    font-size: 13px;
  }

  /* Weight hint next to scores */
  .weight-hint {
    font-size: 11px;
    color: #8b949e;
    font-weight: 400;
    margin-left: 2px;
  }

  /* Scenario badge / button */
  .scenario-btn {
    background: #ddf4ff;
    color: #0969da;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 500;
    border: none;
    cursor: pointer;
    transition: background 0.15s;
    position: relative;
  }
  .scenario-btn:hover { background: #b6e3ff; }
  .scenario-btn .blend-tip {
    display: none;
    position: absolute;
    bottom: 100%;
    left: 50%;
    transform: translateX(-50%);
    background: var(--color-fg);
    color: #fff;
    padding: 6px 10px;
    border-radius: 4px;
    font-size: 11px;
    white-space: nowrap;
    margin-bottom: 6px;
    z-index: 10;
  }
  .scenario-btn:hover .blend-tip { display: block; }
  .scenario-badge {
    background: #ddf4ff;
    color: #0969da;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 500;
  }

  /* Weights dialog */
  dialog {
    border: 1px solid var(--color-border);
    border-radius: var(--radius);
    padding: 0;
    max-width: 420px;
    width: 90%;
    box-shadow: 0 8px 30px rgba(0,0,0,0.12);
  }
  dialog::backdrop { background: rgba(27,31,36,0.5); }
  .dialog-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 16px 20px;
    border-bottom: 1px solid var(--color-border-muted);
  }
  .dialog-header h3 { margin: 0; font-size: 16px; font-weight: 600; }
  .dialog-close {
    background: none;
    border: 1px solid var(--color-border);
    border-radius: var(--radius);
    padding: 4px 12px;
    cursor: pointer;
    font-size: 13px;
    color: var(--color-fg-muted);
  }
  .dialog-close:hover { background: var(--color-bg-subtle); }
  .dialog-body { padding: 16px 20px; }
  .weights-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }
  .weights-table th, .weights-table td {
    padding: 6px 10px;
    border-bottom: 1px solid var(--color-border-muted);
  }
  .weights-table th { background: var(--color-bg-subtle); text-align: left; font-weight: 600; }
  .weights-table td:last-child { text-align: right; font-weight: 600; }
  .weights-table tr:last-child td { border-bottom: none; }

  /* ── Case Detail View ──────────────────────────────────── */
  .card {
    background: var(--color-bg);
    border: 1px solid var(--color-border);
    border-radius: var(--radius);
    margin-bottom: 16px;
    overflow: hidden;
  }
  .card-header {
    padding: 16px 20px;
    border-bottom: 1px solid var(--color-border-muted);
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px;
  }
  .card-header h2 { margin: 0; font-size: 22px; font-weight: 600; }
  .card-body { padding: 20px; }

  /* Badges */
  .badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    color: #fff;
  }
  .badge-score { min-width: 28px; text-align: center; }
  .score-red { background-color: var(--color-red); }
  .score-orange { background-color: var(--color-orange); }
  .score-green { background-color: var(--color-green); }
  .score-dark-green { background-color: var(--color-dark-green); }
  .score-na { background-color: #8b949e; }

  /* Metadata grid */
  .metadata-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
    gap: 8px 24px;
    margin-bottom: 16px;
  }
  .metadata-item { font-size: 13px; }
  .metadata-item .meta-label {
    font-weight: 600;
    color: var(--color-fg-muted);
    text-transform: uppercase;
    font-size: 11px;
    letter-spacing: 0.3px;
  }
  .metadata-item .meta-value { color: var(--color-fg); }

  /* Prompt block */
  .prompt-wrapper { position: relative; margin-bottom: 16px; }
  .prompt-block {
    background: var(--color-bg-subtle);
    border: 1px solid var(--color-border-muted);
    border-radius: var(--radius);
    padding: 14px 18px;
    font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
    font-size: 13px;
    line-height: 1.6;
    white-space: pre-wrap;
    word-break: break-word;
    max-height: 300px;
    overflow-y: auto;
  }
  .copy-btn {
    position: absolute;
    top: 8px;
    right: 8px;
    background: var(--color-bg);
    border: 1px solid var(--color-border);
    border-radius: 4px;
    padding: 3px 10px;
    font-size: 11px;
    cursor: pointer;
    color: var(--color-fg-muted);
    opacity: 0.7;
    transition: opacity 0.15s;
  }
  .copy-btn:hover { opacity: 1; background: var(--color-bg-subtle); }

  /* File path */
  .file-path {
    font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
    font-size: 12px;
    background: var(--color-bg-subtle);
    padding: 2px 6px;
    border-radius: 3px;
    border: 1px solid var(--color-border-muted);
  }
  .file-list { list-style: none; padding: 0; margin: 0 0 16px 0; }
  .file-list li {
    padding: 4px 0;
    font-size: 13px;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .file-desc { color: var(--color-fg-muted); font-size: 12px; }

  /* Section headings inside detail */
  .section-title {
    font-size: 15px;
    font-weight: 600;
    margin: 20px 0 10px 0;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--color-border-muted);
  }
  .section-title:first-child { margin-top: 0; }

  /* Evidence and feedback sentiment coloring */
  .evidence-good { color: var(--color-green); }
  .evidence-bad { color: var(--color-red); }
  .evidence-unconfirmed { color: var(--color-orange); text-decoration: line-through; opacity: 0.7; }
  .feedback-list { margin: 0; padding-left: 18px; line-height: 1.7; }
  .feedback-list li { margin-bottom: 4px; }
  .feedback-list li.evidence-good { color: var(--color-green); }
  .feedback-list li.evidence-bad { color: var(--color-red); }

  /* Score bar */
  .score-bar-row {
    display: flex;
    align-items: center;
    margin-bottom: 8px;
    gap: 10px;
  }
  .score-bar-label {
    width: 180px;
    font-size: 13px;
    flex-shrink: 0;
    text-align: right;
    color: var(--color-fg-muted);
  }
  .score-bar-track {
    flex: 1;
    height: 14px;
    background: #eaeef2;
    border-radius: 7px;
    overflow: hidden;
  }
  .score-bar-fill {
    height: 100%;
    border-radius: 7px;
    transition: width 0.3s;
  }
  .score-bar-value {
    width: 40px;
    font-size: 13px;
    font-weight: 600;
    text-align: right;
    flex-shrink: 0;
  }

  /* Block averages */
  .block-avgs {
    display: flex;
    gap: 16px;
    margin-top: 12px;
    flex-wrap: wrap;
  }
  .block-avg-chip {
    padding: 6px 14px;
    border-radius: var(--radius);
    font-size: 13px;
    font-weight: 500;
    background: var(--color-bg-subtle);
    border: 1px solid var(--color-border-muted);
  }
  .block-avg-chip strong { font-weight: 600; }

  /* Collapsible details */
  details {
    border: 1px solid var(--color-border-muted);
    border-radius: var(--radius);
    margin-bottom: 8px;
  }
  details summary {
    padding: 10px 16px;
    cursor: pointer;
    font-weight: 500;
    font-size: 13px;
    display: flex;
    align-items: center;
    gap: 10px;
    user-select: none;
    list-style: none;
  }
  details summary::-webkit-details-marker { display: none; }
  details summary::before {
    content: "\25B6";
    font-size: 10px;
    color: var(--color-fg-muted);
    transition: transform 0.15s;
  }
  details[open] summary::before { transform: rotate(90deg); }
  details .detail-body {
    padding: 12px 16px 16px 16px;
    border-top: 1px solid var(--color-border-muted);
  }
  .dim-grayed summary { color: var(--color-fg-muted); }
  .dim-grayed .detail-body { color: var(--color-fg-muted); }

  .feedback { margin: 0 0 10px 0; line-height: 1.6; }
  .feedback-list { margin: 0; padding-left: 18px; line-height: 1.7; list-style: none; }
  .feedback-list li { margin-bottom: 4px; }
  .evidence-list { margin: 0; padding-left: 18px; list-style: none; }
  .evidence-list li { margin-bottom: 4px; font-size: 13px; color: var(--color-fg-muted); }

  /* Responsive */
  @media (max-width: 640px) {
    body { padding: 12px; }
    .header { padding: 16px; }
    .score-bar-label { width: 100px; font-size: 12px; }
    .metadata-grid { grid-template-columns: 1fr; }
  }
</style>
</head>
<body>
<div class="container">

  <!-- Shared Header -->
  <div class="header">
    <div>
      <h1>Excel Gen Eval Report</h1>
      <div class="header-meta">Generated {{ timestamp }} &bull; {{ results | length }} case{{ "s" if results | length != 1 else "" }}</div>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════════════════════
       VIEW 1: Summary Table
       ═══════════════════════════════════════════════════════════ -->
  <div id="view-summary">
    <div class="summary-wrapper">
      <h2>Summary</h2>
      <table class="summary-table">
        <thead>
          <tr>
            <th rowspan="2" style="text-align:center;width:30px;">#</th>
            <th rowspan="2" style="text-align:left;">Case ID</th>
            <th rowspan="2" style="text-align:left;">Task ID</th>
            <th rowspan="2" style="text-align:left;">Scenario</th>
            <th colspan="4" class="block-header">Data &amp; Content</th>
            <th colspan="4" class="block-header">Structure &amp; Usability</th>
            <th rowspan="2">Weighted Avg</th>
          </tr>
          <tr>
            {% for dim_key in all_dims %}
            <th title="{{ dim_labels[dim_key] }}">{{ dim_short[dim_key] }}</th>
            {% endfor %}
          </tr>
        </thead>
        <tbody>
          {% for r in results %}
          {% set blended_w = get_blended_weights(r.scenario) %}
          <tr>
            <td style="text-align:center;color:var(--color-fg-muted);">{{ loop.index }}</td>
            <td><a class="case-link" href="javascript:void(0)" onclick="showCase({{ loop.index0 }})">{{ r.case_id }}</a></td>
            <td style="font-size:11px;color:var(--color-fg-muted);font-family:monospace;">{{ r.metadata.get('task_id', '') if r.metadata else '' }}</td>
            <td>
              <button class="scenario-btn" type="button"
                      onclick="document.getElementById('dlg-{{ loop.index0 }}').showModal()">
                {{ r.scenario.detected.value | replace("_", " ") | title }}
                {% if r.scenario.blend | length > 1 %}
                <span class="blend-tip">{% for sk, sv in r.scenario.blend.items() %}{{ sk | replace("_", " ") | title }}: {{ "%.0f" | format(sv * 100) }}%{% if not loop.last %} · {% endif %}{% endfor %}</span>
                {% endif %}
              </button>
            </td>
            {% for dim_key in all_dims %}
            {% set dr = r.dimensions.get(dim_key) %}
            {% set w = blended_w[dim_key] %}
            <td>
              {% if dr and dr.score is not none %}
              <span class="score-cell {{ score_color(dr.score) }}">{{ dr.score }}</span>{% if w != 1.0 %}<span class="weight-hint">&times;{{ "%.1f" | format(w) }}</span>{% endif %}
              {% elif dr and dr.status.value == "na" %}
              <span style="color:var(--color-fg-muted);">N/A</span>
              {% elif dr and dr.status.value == "skipped" %}
              <span style="color:var(--color-fg-muted);">Skip</span>
              {% else %}
              <span style="color:var(--color-fg-muted);">&mdash;</span>
              {% endif %}
            </td>
            {% endfor %}
            <td>
              {% if r.summary.overall_weighted_avg is not none %}
              <span class="score-cell {{ score_color(r.summary.overall_weighted_avg | round(0) | int) }}">{{ "%.2f" | format(r.summary.overall_weighted_avg) }}</span>
              {% else %}&mdash;{% endif %}
            </td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>{# /view-summary #}

  <!-- Methodology (collapsible, inside summary view) -->
  <div id="view-methodology" class="card" style="margin-bottom:24px;">
    <div class="card-header">
      <h2>Scoring Methodology &amp; Rules</h2>
    </div>
    <div class="card-body" style="font-size:13px;line-height:1.7;">
        <h4 style="margin:0 0 8px;">Evaluation Approach</h4>
        <p>Each Excel workbook is evaluated by an LLM judge across <strong>8 dimensions</strong> in two blocks. A <strong>code-level data scanner</strong> pre-computes factual metrics (formula errors, column statistics, source-vs-generated row diffs) that the LLM uses as ground truth.</p>

        <h4 style="margin:16px 0 8px;">Scoring Scale (0&ndash;4)</h4>
        <table class="weights-table">
          <thead><tr><th>Score</th><th>Meaning</th></tr></thead>
          <tbody>
            <tr><td><span class="score-cell score-red">0</span></td><td>Unacceptable &mdash; major failures, unusable</td></tr>
            <tr><td><span class="score-cell score-red">1</span></td><td>Poor &mdash; significant problems, heavy rework needed</td></tr>
            <tr><td><span class="score-cell score-orange">2</span></td><td>Adequate &mdash; meets basic needs, lacks depth</td></tr>
            <tr><td><span class="score-cell score-green">3</span></td><td>Good &mdash; solid and professional, minor issues</td></tr>
            <tr><td><span class="score-cell score-dark-green">4</span></td><td>Exceptional &mdash; flawless, publication-ready</td></tr>
          </tbody>
        </table>

        <h4 style="margin:16px 0 8px;">Dimensions</h4>
        <table class="weights-table">
          <thead><tr><th>Block</th><th>Dimension</th><th>What It Measures</th></tr></thead>
          <tbody>
            <tr><td rowspan="4"><strong>Data &amp; Content</strong></td><td>Data Accuracy</td><td>Output values correct and consistent with source data; zero tolerance for fabrication</td></tr>
            <tr><td>Completeness</td><td>All requested data points, metrics, and analyses present</td></tr>
            <tr><td>Formula &amp; Logic</td><td>Formulas used (vs hardcoded), well-structured, error-free</td></tr>
            <tr><td>Relevance</td><td>Content stays focused on user&rsquo;s request</td></tr>
            <tr><td rowspan="4"><strong>Structure &amp; Usability</strong></td><td>Sheet Organization</td><td>Multi-sheet structure, naming, navigation</td></tr>
            <tr><td>Table Structure</td><td>Headers, data types, merged cells, Excel Table usage</td></tr>
            <tr><td>Chart Appropriateness</td><td>Chart type selection, labeling, data communication</td></tr>
            <tr><td>Professional Formatting</td><td>Colors, fonts, conditional formatting, visual polish</td></tr>
          </tbody>
        </table>

        <h4 style="margin:16px 0 8px;">Scenario-Based Weighting</h4>
        <p>The tool auto-detects the workbook scenario and applies scenario-specific dimension weights. Different scenarios emphasize different quality aspects (e.g., financial modeling weights Formula &amp; Logic heavily, while reporting weights Charts heavily).</p>

        <p style="margin-top:8px;"><strong>Scenario Weight Table</strong></p>
        <div style="overflow-x:auto;">
        <table class="weights-table" style="font-size:12px;">
          <thead>
            <tr><th>Scenario</th>{% for dim_key in all_dims %}<th>{{ dim_short[dim_key] }}</th>{% endfor %}</tr>
          </thead>
          <tbody>
            {% for scen, scen_label in [("reporting_analysis","Reporting Analysis"),("data_processing","Data Processing"),("template_form","Template Form"),("planning_tracking","Planning Tracking"),("financial_modeling","Financial Modeling"),("comparison_evaluation","Comparison Evaluation"),("general","General")] %}
            <tr>
              <td style="white-space:nowrap;">{{ scen_label }}</td>
              {% for dim_key in all_dims %}
              {% set w = scenario_weights[scen][dim_key] if scen in scenario_weights else 1.0 %}
              <td style="text-align:center;{% if w > 1.0 %}font-weight:700;color:var(--color-green);{% elif w < 1.0 %}color:var(--color-fg-muted);{% endif %}">{{ "%.1f" | format(w) }}</td>
              {% endfor %}
            </tr>
            {% endfor %}
          </tbody>
        </table>
        </div>

        <h4 style="margin:16px 0 8px;">Multi-Scenario Blending</h4>
        <p>Many workbooks span multiple scenarios (e.g., an audit report that also involves data processing). Instead of forcing a single classification, the tool detects a <strong>blend</strong> of up to 2 scenarios with proportional weights.</p>
        <p style="margin-top:4px;">Example: A workbook classified as <em>70% Reporting + 30% Data Processing</em>:</p>
        <table class="weights-table" style="font-size:12px;max-width:500px;">
          <thead><tr><th>Dimension</th><th>Reporting (&times;0.7)</th><th>Data Proc (&times;0.3)</th><th>Blended</th></tr></thead>
          <tbody>
            <tr><td>Data Accuracy</td><td>1.0 &times; 0.7 = 0.70</td><td>1.5 &times; 0.3 = 0.45</td><td><strong>1.15</strong></td></tr>
            <tr><td>Prof Formatting</td><td>0.8 &times; 0.7 = 0.56</td><td>0.5 &times; 0.3 = 0.15</td><td><strong>0.71</strong></td></tr>
          </tbody>
        </table>
        <p style="margin-top:4px;font-size:12px;color:var(--color-fg-muted);">Formula: <code>blended_weight = &sum;(scenario_weight &times; scenario_proportion)</code>. Click a scenario badge in the summary table to see the actual blended weights for that case.</p>

        <h4 style="margin:16px 0 8px;">Evidence Tags</h4>
        <ul style="padding-left:18px;">
          <li><span class="evidence-good">+VERIFIED</span> &mdash; positive finding, independently confirmed</li>
          <li><span class="evidence-bad">-VERIFIED</span> &mdash; negative finding, independently confirmed</li>
          <li><span class="evidence-good">+INFERRED</span> &mdash; positive observation based on patterns</li>
          <li><span class="evidence-bad">-INFERRED</span> &mdash; concern based on patterns, not fully verified</li>
          <li><span class="evidence-unconfirmed">UNCONFIRMED</span> &mdash; claim could not be confirmed by independent verification (strikethrough)</li>
        </ul>

        <h4 style="margin:16px 0 8px;">Anti-Hallucination Measures</h4>
        <p>A code-level data scanner pre-computes sheet profiles, formula error counts, and row-by-row source-vs-generated diffs. These factual findings are injected into LLM prompts as ground truth, reducing hallucinated evidence. LLM judges are instructed to only cite values they can directly verify.</p>
      </div>
  </div>

  <!-- Scenario Weights Dialogs (shared, always in DOM) -->
  {% for r in results %}
  {% set blended_w = get_blended_weights(r.scenario) %}
  <dialog id="dlg-{{ loop.index0 }}">
    <div class="dialog-header">
      <h3>Scenario Weights — {{ r.case_id }}</h3>
      <button class="dialog-close" type="button"
              onclick="document.getElementById('dlg-{{ loop.index0 }}').close()">Close</button>
    </div>
    <div class="dialog-body">
      {% if r.scenario.blend | length > 1 %}
      <p style="font-size:13px;color:var(--color-fg-muted);margin:0 0 12px;">
        Blended scenario: {% for sk, sv in r.scenario.blend.items() %}<strong>{{ sk | replace("_", " ") | title }}</strong> {{ "%.0f" | format(sv * 100) }}%{% if not loop.last %} + {% endif %}{% endfor %}
      </p>
      {% else %}
      <p style="font-size:13px;color:var(--color-fg-muted);margin:0 0 12px;">
        Scenario: <strong>{{ r.scenario.detected.value | replace("_", " ") | title }}</strong> ({{ "%.0f" | format(r.scenario.confidence * 100) }}% confidence)
      </p>
      {% endif %}
      <table class="weights-table">
        <thead><tr><th>Dimension</th><th>Blended Weight</th></tr></thead>
        <tbody>
          {% for dim_key in all_dims %}
          <tr>
            <td>{{ dim_labels[dim_key] }}</td>
            <td>{{ "%.2f" | format(blended_w[dim_key]) }}</td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </dialog>
  {% endfor %}

  <!-- ═══════════════════════════════════════════════════════════
       VIEW 2: Case Detail (one div per case, hidden by default)
       ═══════════════════════════════════════════════════════════ -->
  {% for r in results %}
  {% set weights = get_blended_weights(r.scenario) %}
  <div id="case-detail-{{ loop.index0 }}" class="view-hidden">

    <a class="back-link" href="javascript:void(0)" onclick="showSummary()">&larr; Back to Summary</a>

    <!-- Header card -->
    <div class="card">
      <div class="card-header">
        <h2>{{ r.case_id }}</h2>
        <div style="display:flex;align-items:center;gap:10px;">
          <button class="scenario-btn" type="button"
                  onclick="document.getElementById('dlg-{{ loop.index0 }}').showModal()">
            {{ r.scenario.detected.value | replace("_", " ") | title }}
          </button>
          {% if r.summary.overall_weighted_avg is not none %}
          <span class="badge badge-score {{ score_color(r.summary.overall_weighted_avg | round(0) | int) }}">{{ "%.2f" | format(r.summary.overall_weighted_avg) }}</span>
          {% endif %}
        </div>
      </div>
      <div class="card-body">

        {# ── Metadata ────────────────────────────────────── #}
        {% set meta_items = [] %}
        {% for mk in metadata_keys %}
          {% if r.metadata.get(mk) is not none %}
            {% if meta_items.append((mk, r.metadata[mk])) %}{% endif %}
          {% endif %}
        {% endfor %}
        {% if meta_items %}
        <div class="section-title">Metadata</div>
        <div class="metadata-grid">
          {% for mk, mv in meta_items %}
          <div class="metadata-item">
            <div class="meta-label">{{ mk | replace("_", " ") }}</div>
            <div class="meta-value">{{ mv }}</div>
          </div>
          {% endfor %}
        </div>
        {% endif %}

        {# ── Prompt ──────────────────────────────────────── #}
        {% if r.prompt %}
        <div class="section-title">Prompt</div>
        <div class="prompt-wrapper">
          <button class="copy-btn" onclick="navigator.clipboard.writeText(this.nextElementSibling.textContent).then(()=>{this.textContent='Copied!';setTimeout(()=>this.textContent='Copy',1500)})">Copy</button>
          <div class="prompt-block">{{ r.prompt }}</div>
        </div>
        {% endif %}

        {# ── Files ───────────────────────────────────────── #}
        {% if r.input_files or r.output_files %}
        <div class="section-title">Files</div>
        {% if r.input_files %}
        <div style="font-size:13px;font-weight:600;margin-bottom:4px;">Input Files</div>
        <ul class="file-list">
          {% for f in r.input_files %}
          <li><span class="file-path">{{ f.path }}</span>{% if f.description %}<span class="file-desc">{{ f.description }}</span>{% endif %}</li>
          {% endfor %}
        </ul>
        {% endif %}
        {% if r.output_files %}
        <div style="font-size:13px;font-weight:600;margin-bottom:4px;">Output Files</div>
        <ul class="file-list">
          {% for f in r.output_files %}
          <li><span class="file-path">{{ f.path }}</span>{% if f.description %}<span class="file-desc">{{ f.description }}</span>{% endif %}</li>
          {% endfor %}
        </ul>
        {% endif %}
        {% endif %}

        {# ── Score Overview ──────────────────────────────── #}
        <div class="section-title">Score Overview</div>
        {% for dim_key in all_dims %}
        {% set dr = r.dimensions.get(dim_key) %}
        <div class="score-bar-row">
          <span class="score-bar-label">{{ dim_labels[dim_key] }}</span>
          <div class="score-bar-track">
            {% if dr and dr.score is not none %}
            <div class="score-bar-fill {{ score_color(dr.score) }}" style="width:{{ score_pct(dr.score) }}%"></div>
            {% endif %}
          </div>
          <span class="score-bar-value">
            {% if dr and dr.score is not none %}{{ dr.score }}/4{% elif dr and dr.status.value == "na" %}N/A{% elif dr and dr.status.value == "skipped" %}Skip{% else %}&mdash;{% endif %}
          </span>
        </div>
        {% endfor %}

        <div class="block-avgs">
          <div class="block-avg-chip">Data &amp; Content avg: <strong>{% if r.summary.data_content_avg is not none %}{{ "%.2f" | format(r.summary.data_content_avg) }}{% else %}&mdash;{% endif %}</strong></div>
          <div class="block-avg-chip">Structure &amp; Usability avg: <strong>{% if r.summary.structure_usability_avg is not none %}{{ "%.2f" | format(r.summary.structure_usability_avg) }}{% else %}&mdash;{% endif %}</strong></div>
        </div>

        {# ── Dimension Details ───────────────────────────── #}
        <div class="section-title" style="margin-top:24px;">Dimension Details</div>
        {% for dim_key in all_dims %}
        {% set dr = r.dimensions.get(dim_key) %}
        {% set w = weights[dim_key] %}
        {% set is_inactive = (not dr) or dr.status.value in ("na", "skipped") %}
        <details open class="{{ 'dim-grayed' if is_inactive else '' }}">
          <summary>
            {% if dr and dr.score is not none %}
            <span class="badge badge-score {{ score_color(dr.score) }}">{{ dr.score }}</span>
            {% elif dr and dr.status.value == "na" %}
            <span class="badge score-na">N/A</span>
            {% elif dr and dr.status.value == "skipped" %}
            <span class="badge score-na">Skipped</span>
            {% else %}
            <span class="badge score-na">&mdash;</span>
            {% endif %}
            {{ dim_labels[dim_key] }}
            {% if w != 1.0 %}<span class="weight-hint">(&times;{{ "%.1f" | format(w) }})</span>{% endif %}
          </summary>
          <div class="detail-body">
            {% if dr %}
              {% set fb = dr.feedback_zh if dr.feedback_zh else dr.feedback %}
              {% if fb %}{{ highlight_feedback(fb) }}{% endif %}
              {% set ev_list = dr.evidence_zh if dr.evidence_zh else dr.evidence %}
              {% if ev_list %}
              <p style="font-size:12px;font-weight:600;color:var(--color-fg-muted);margin:12px 0 4px;">Key Findings</p>
              <ul class="evidence-list">
                {% for ev in ev_list %}<li>{{ highlight_evidence(ev) }}</li>{% endfor %}
              </ul>
              {% endif %}
              {% if dr.error_message %}<p style="color:var(--color-red)">Error: {{ dr.error_message }}</p>{% endif %}
              {% if is_inactive %}<p style="font-style:italic;">This dimension was {{ dr.status.value }}.</p>{% endif %}
            {% else %}
              <p style="font-style:italic;">No data available for this dimension.</p>
            {% endif %}
          </div>
        </details>
        {% endfor %}

      </div>{# /card-body #}
    </div>{# /card #}
  </div>{# /case-detail-N #}
  {% endfor %}

</div>{# /container #}

<script>
(function() {
  var caseCount = {{ results | length }};

  function showCase(index) {
    document.getElementById('view-summary').classList.add('view-hidden');
    document.getElementById('view-methodology').classList.add('view-hidden');
    for (var i = 0; i < caseCount; i++) {
      document.getElementById('case-detail-' + i).classList.add('view-hidden');
    }
    document.getElementById('case-detail-' + index).classList.remove('view-hidden');
    window.scrollTo(0, 0);
  }

  function showSummary() {
    for (var i = 0; i < caseCount; i++) {
      document.getElementById('case-detail-' + i).classList.add('view-hidden');
    }
    document.getElementById('view-summary').classList.remove('view-hidden');
    document.getElementById('view-methodology').classList.remove('view-hidden');
    window.scrollTo(0, 0);
  }

  // Expose globally
  window.showCase = showCase;
  window.showSummary = showSummary;
})();
</script>
</body>
</html>
"""


def generate_html_report(
    results: list[EvalResult],
    output_path: str | Path,
) -> Path:
    """Render evaluation results to a self-contained HTML file.

    The report contains two views toggled via JavaScript:
      1. A summary table with all cases and scores.
      2. A detail view per case with metadata, prompt, files,
         score bars, and highlighted dimension feedback.

    Parameters
    ----------
    results:
        One or more ``EvalResult`` objects to include in the report.
    output_path:
        Destination file path for the generated HTML.

    Returns
    -------
    Path
        The resolved path to the written HTML file.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    env = Environment(autoescape=True)
    # Register highlight_feedback so it returns raw (safe) HTML
    env.globals["highlight_feedback"] = lambda text: Markup(_highlight_feedback(text))
    env.globals["highlight_evidence"] = lambda text: Markup(_highlight_evidence(text))

    template = env.from_string(_HTML_TEMPLATE)

    html = template.render(
        results=results,
        timestamp=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        all_dims=_ALL_DIMS,
        dim_labels=_DIMENSION_LABELS,
        dim_short=_DIMENSION_SHORT,
        score_color=_score_color,
        score_pct=_score_pct,
        scenario_weights={s.value: w for s, w in SCENARIO_WEIGHTS.items()},
        get_blended_weights=get_blended_weights,
        metadata_keys=_METADATA_KEYS,
    )

    output_path.write_text(html, encoding="utf-8")
    return output_path.resolve()
