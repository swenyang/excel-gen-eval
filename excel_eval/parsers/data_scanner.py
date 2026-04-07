"""Code-level data scanner — produces objective scan reports for LLM evaluators.

Scans generated Excel data and (optionally) source data to produce factual
observations that the LLM judge can rely on without hallucinating.
"""

from __future__ import annotations

import io
import math
from dataclasses import dataclass, field
from difflib import SequenceMatcher

import pandas as pd


@dataclass
class ColumnProfile:
    """Statistical profile of a single column."""
    name: str
    dtype: str  # "numeric", "string", "mixed", "empty"
    non_null_count: int
    null_count: int
    unique_count: int
    sample_values: list[str] = field(default_factory=list)  # up to 5
    # Numeric stats (if applicable)
    num_min: float | None = None
    num_max: float | None = None
    num_mean: float | None = None
    num_sum: float | None = None


@dataclass
class SheetProfile:
    """Profile of a single sheet."""
    name: str
    row_count: int
    col_count: int
    columns: list[ColumnProfile] = field(default_factory=list)
    empty_rows: int = 0
    duplicate_rows: int = 0


@dataclass
class ColumnMapping:
    """Mapping between a source column and a generated column."""
    source_col: str
    generated_col: str
    match_type: str  # "exact_name", "fuzzy_name", "value_overlap"
    confidence: float  # 0.0-1.0
    value_match_pct: float | None = None  # % of values that match


@dataclass
class RowDiff:
    """A single row difference between source and generated data."""
    row_index: int
    column: str
    source_value: str
    generated_value: str


@dataclass
class DataComparisonReport:
    """Comparison between source and generated data."""
    source_rows: int
    generated_rows: int
    column_mappings: list[ColumnMapping] = field(default_factory=list)
    matched_rows: int = 0
    total_comparable_rows: int = 0
    match_percentage: float = 0.0
    row_diffs: list[RowDiff] = field(default_factory=list)  # capped at 50
    summary: str = ""


@dataclass
class ScanReport:
    """Complete scan report for a generated workbook."""
    sheet_profiles: list[SheetProfile] = field(default_factory=list)
    formula_error_count: int = 0
    formula_error_details: list[str] = field(default_factory=list)
    total_formulas: int = 0
    hardcoded_calculation_suspects: int = 0
    data_comparison: DataComparisonReport | None = None


def scan_generated_excel(
    generated_csv_texts: dict[str, str],
    source_text: str | None = None,
    source_dataframes: dict[str, "pd.DataFrame"] | None = None,
    formulas: list | None = None,
) -> ScanReport:
    """Scan generated Excel data and produce an objective report.

    Args:
        generated_csv_texts: {sheet_name: csv_text} from the generated workbook
        source_text: raw text of the source/grounding data (any format)
        source_dataframes: {sheet_name: DataFrame} from the source Excel (preferred for comparison)
        formulas: list of FormulaInfo objects from excel_parser
    """
    report = ScanReport()

    # Profile each sheet
    for sheet_name, csv_text in generated_csv_texts.items():
        profile = _profile_sheet(sheet_name, csv_text)
        report.sheet_profiles.append(profile)

    # Analyze formulas
    if formulas:
        report.total_formulas = len(formulas)
        for f in formulas:
            if f.has_error:
                report.formula_error_count += 1
                report.formula_error_details.append(
                    f"{f.sheet}!{f.cell}: {f.formula} → {f.computed_value}"
                )

    # Compare with source data — match multiple sheets when possible
    comparisons: list[DataComparisonReport] = []

    if source_dataframes and generated_csv_texts:
        gen_dfs: dict[str, pd.DataFrame] = {}
        for name, csv_text in generated_csv_texts.items():
            df = _try_parse_as_dataframe(csv_text)
            if df is not None and len(df) > 0:
                gen_dfs[name] = df

        used_gen: set[str] = set()
        for src_name, src_df in source_dataframes.items():
            if len(src_df) == 0:
                continue
            best_name, best_score = None, 0.0
            for gen_name, gen_df in gen_dfs.items():
                if gen_name in used_gen:
                    continue
                name_sim = SequenceMatcher(None, src_name.lower(), gen_name.lower()).ratio()
                col_overlap = len(set(str(c) for c in src_df.columns) & set(str(c) for c in gen_df.columns))
                col_total = max(len(set(str(c) for c in src_df.columns) | set(str(c) for c in gen_df.columns)), 1)
                score = name_sim * 0.4 + (col_overlap / col_total) * 0.6
                if score > best_score:
                    best_score = score
                    best_name = gen_name
            if best_name and best_score >= 0.3:
                used_gen.add(best_name)
                comp = _compare_dataframes(src_df, gen_dfs[best_name])
                comp.summary = f"[{src_name} → {best_name}] {comp.summary}"
                comparisons.append(comp)
    elif source_text and generated_csv_texts:
        source_df = _try_parse_as_dataframe(source_text)
        if source_df is not None:
            largest_sheet = max(generated_csv_texts.items(), key=lambda x: len(x[1]))
            gen_df = _try_parse_as_dataframe(largest_sheet[1])
            if gen_df is not None:
                comparisons.append(_compare_dataframes(source_df, gen_df))

    if comparisons:
        report.data_comparison = comparisons[0]

    return report


def format_scan_report(report: ScanReport) -> str:
    """Format a ScanReport as human-readable text for LLM consumption."""
    lines: list[str] = []

    lines.append("## Code-Level Scan Report (Automated — Factual)")
    lines.append("*The following observations are computed by code, not by AI. Treat them as ground truth.*\n")

    # Sheet profiles
    lines.append("### Sheet Profiles")
    for sp in report.sheet_profiles:
        lines.append(f"**{sp.name}**: {sp.row_count} rows × {sp.col_count} cols "
                      f"({sp.empty_rows} empty rows, {sp.duplicate_rows} duplicate rows)")
        for cp in sp.columns:
            stats = f"  - `{cp.name}` ({cp.dtype}): {cp.non_null_count} non-null, {cp.null_count} null, {cp.unique_count} unique"
            if cp.num_mean is not None:
                stats += f", mean={cp.num_mean:.2f}, sum={cp.num_sum:.2f}"
            if cp.sample_values:
                samples = ", ".join(cp.sample_values[:3])
                stats += f", samples=[{samples}]"
            lines.append(stats)

    # Formulas
    lines.append(f"\n### Formula Analysis")
    lines.append(f"- Total formulas: {report.total_formulas}")
    lines.append(f"- Formula errors: {report.formula_error_count}")
    if report.formula_error_details:
        lines.append("- Error details:")
        for err in report.formula_error_details[:20]:
            lines.append(f"  - {err}")

    # Data comparison
    if report.data_comparison:
        dc = report.data_comparison
        lines.append(f"\n### Source vs Generated Data Comparison")
        lines.append(f"- Source rows: {dc.source_rows}, Generated rows: {dc.generated_rows}")
        lines.append(f"- Comparable rows: {dc.total_comparable_rows}, Matched: {dc.matched_rows} ({dc.match_percentage:.1f}%)")

        if dc.column_mappings:
            lines.append("- Column mappings:")
            for cm in dc.column_mappings:
                conf = f"{cm.confidence:.0%}"
                match_info = f"value_match={cm.value_match_pct:.1f}%" if cm.value_match_pct is not None else ""
                lines.append(f"  - `{cm.source_col}` → `{cm.generated_col}` ({cm.match_type}, conf={conf}) {match_info}")

        if dc.row_diffs:
            lines.append(f"- Differences found ({len(dc.row_diffs)} shown, may be more):")
            for rd in dc.row_diffs[:30]:
                lines.append(f"  - Row {rd.row_index}, `{rd.column}`: source=`{rd.source_value}` vs generated=`{rd.generated_value}`")

        if dc.summary:
            lines.append(f"- Summary: {dc.summary}")
    else:
        lines.append("\n### Source vs Generated Data Comparison")
        lines.append("- No structured comparison possible (source data format not tabular or not provided)")

    return "\n".join(lines)


# ── Internal helpers ───────────────────────────────────────────────────────


def _profile_sheet(name: str, csv_text: str) -> SheetProfile:
    """Profile a single sheet from its CSV text."""
    df = _try_parse_as_dataframe(csv_text)
    if df is None:
        return SheetProfile(name=name, row_count=0, col_count=0)

    columns = []
    for col_name in df.columns:
        col = df[col_name]
        non_null = int(col.notna().sum())
        null_count = int(col.isna().sum())
        unique = int(col.nunique())

        # Determine dtype
        if non_null == 0:
            dtype = "empty"
        else:
            numeric_count = pd.to_numeric(col, errors="coerce").notna().sum()
            if numeric_count > non_null * 0.8:
                dtype = "numeric"
            elif numeric_count < non_null * 0.2:
                dtype = "string"
            else:
                dtype = "mixed"

        # Compute numeric stats
        num_min = num_max = num_mean = num_sum = None
        if dtype == "numeric":
            num_col = pd.to_numeric(col, errors="coerce").dropna()
            if len(num_col) > 0:
                num_min = float(num_col.min())
                num_max = float(num_col.max())
                num_mean = float(num_col.mean())
                num_sum = float(num_col.sum())

        # Sample values
        samples = [str(v) for v in col.dropna().head(5).tolist()]

        columns.append(ColumnProfile(
            name=str(col_name),
            dtype=dtype,
            non_null_count=non_null,
            null_count=null_count,
            unique_count=unique,
            sample_values=samples,
            num_min=num_min, num_max=num_max,
            num_mean=num_mean, num_sum=num_sum,
        ))

    empty_rows = int((df.isna().all(axis=1) | (df == "")).all(axis=1).sum())
    duplicate_rows = int(df.duplicated().sum())

    return SheetProfile(
        name=name,
        row_count=len(df),
        col_count=len(df.columns),
        columns=columns,
        empty_rows=empty_rows,
        duplicate_rows=duplicate_rows,
    )


def _compare_dataframes(source_df: pd.DataFrame, gen_df: pd.DataFrame) -> DataComparisonReport:
    """Compare source and generated DataFrames."""
    report = DataComparisonReport(
        source_rows=len(source_df),
        generated_rows=len(gen_df),
    )

    # Step 1: Match columns
    col_mappings = _match_columns(source_df, gen_df)
    report.column_mappings = col_mappings

    if not col_mappings:
        report.summary = "No column mappings could be established between source and generated data."
        return report

    # Step 2: Compare mapped columns row-by-row
    mapped_src_cols = [cm.source_col for cm in col_mappings if cm.confidence >= 0.5]
    mapped_gen_cols = [cm.generated_col for cm in col_mappings if cm.confidence >= 0.5]

    if not mapped_src_cols:
        report.summary = "Column mappings too low confidence for row comparison."
        return report

    # Use the minimum row count for comparison
    compare_rows = min(len(source_df), len(gen_df))
    report.total_comparable_rows = compare_rows

    matched = 0
    diffs: list[RowDiff] = []

    for i in range(compare_rows):
        row_match = True
        for src_col, gen_col in zip(mapped_src_cols, mapped_gen_cols):
            src_val = source_df.iloc[i].get(src_col)
            gen_val = gen_df.iloc[i].get(gen_col)

            if not _values_match(src_val, gen_val):
                row_match = False
                if len(diffs) < 50:
                    diffs.append(RowDiff(
                        row_index=i + 2,  # 1-indexed + header
                        column=src_col,
                        source_value=str(src_val)[:100],
                        generated_value=str(gen_val)[:100],
                    ))

        if row_match:
            matched += 1

    report.matched_rows = matched
    report.match_percentage = (matched / compare_rows * 100) if compare_rows > 0 else 0
    report.row_diffs = diffs

    diff_count = compare_rows - matched
    report.summary = (
        f"{matched}/{compare_rows} rows match ({report.match_percentage:.1f}%). "
        f"{diff_count} rows have differences across {len(mapped_src_cols)} compared columns."
    )

    # Update value_match_pct on column mappings
    for cm in col_mappings:
        if cm.source_col in mapped_src_cols:
            idx = mapped_src_cols.index(cm.source_col)
            gen_col = mapped_gen_cols[idx]
            match_count = sum(
                1 for i in range(compare_rows)
                if _values_match(source_df.iloc[i].get(cm.source_col), gen_df.iloc[i].get(gen_col))
            )
            cm.value_match_pct = match_count / compare_rows * 100 if compare_rows > 0 else 0

    return report


def _match_columns(source_df: pd.DataFrame, gen_df: pd.DataFrame) -> list[ColumnMapping]:
    """Match source columns to generated columns using name and value similarity."""
    mappings: list[ColumnMapping] = []
    used_gen_cols: set[str] = set()

    for src_col in source_df.columns:
        best_match: ColumnMapping | None = None
        best_score = 0.0

        for gen_col in gen_df.columns:
            if gen_col in used_gen_cols:
                continue

            # Name similarity
            name_sim = SequenceMatcher(None, str(src_col).lower(), str(gen_col).lower()).ratio()

            # Value overlap (for string columns)
            src_vals = set(source_df[src_col].dropna().astype(str).head(100))
            gen_vals = set(gen_df[gen_col].dropna().astype(str).head(100))
            val_overlap = len(src_vals & gen_vals) / max(len(src_vals | gen_vals), 1)

            # Combined score
            score = name_sim * 0.6 + val_overlap * 0.4

            if score > best_score:
                best_score = score
                match_type = "exact_name" if name_sim > 0.95 else ("fuzzy_name" if name_sim > 0.5 else "value_overlap")
                best_match = ColumnMapping(
                    source_col=str(src_col),
                    generated_col=str(gen_col),
                    match_type=match_type,
                    confidence=score,
                )

        if best_match and best_score >= 0.3:
            mappings.append(best_match)
            used_gen_cols.add(best_match.generated_col)

    return mappings


def _values_match(a, b) -> bool:
    """Check if two values match, with tolerance for numeric comparison."""
    if pd.isna(a) and pd.isna(b):
        return True
    if pd.isna(a) or pd.isna(b):
        return False

    # String comparison
    str_a, str_b = str(a).strip(), str(b).strip()
    if str_a == str_b:
        return True

    # Numeric comparison with tolerance
    try:
        num_a, num_b = float(a), float(b)
        if num_a == 0 and num_b == 0:
            return True
        if abs(num_a) < 1e-10:
            return abs(num_b) < 1e-10
        return abs(num_a - num_b) / max(abs(num_a), abs(num_b)) < 0.001  # 0.1% tolerance
    except (ValueError, TypeError):
        return False


def _try_parse_as_dataframe(text: str) -> pd.DataFrame | None:
    """Try to parse text as a CSV DataFrame."""
    try:
        return pd.read_csv(io.StringIO(text))
    except Exception:
        pass

    # Try tab-separated
    try:
        return pd.read_csv(io.StringIO(text), sep="\t")
    except Exception:
        pass

    return None
