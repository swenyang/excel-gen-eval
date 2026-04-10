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
    match_method: str = "positional"  # "positional" or "key_based"
    key_column: str | None = None  # the key column used for key-based matching
    empty_rows_dropped: int = 0  # total empty rows removed from both sides


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
    generated_dataframes: dict[str, "pd.DataFrame"] | None = None,
    formulas: list | None = None,
    hidden_sheets: set[str] | None = None,
) -> ScanReport:
    """Scan generated Excel data and produce an objective report.

    Args:
        generated_csv_texts: {sheet_name: csv_text} for sheet profiling
        source_text: raw text of the source/grounding data (any format)
        source_dataframes: {sheet_name: DataFrame} from the source Excel (raw pandas values)
        generated_dataframes: {sheet_name: DataFrame} from the generated Excel (raw pandas values,
            preferred over CSV for comparison to avoid format mismatches)
        formulas: list of FormulaInfo objects from excel_parser
        hidden_sheets: set of sheet names that are hidden in the generated workbook
    """
    report = ScanReport()

    # Profile each sheet — prefer raw DataFrames (full data) over CSV (may be truncated)
    if generated_dataframes:
        for sheet_name, df in generated_dataframes.items():
            profile = _profile_dataframe(sheet_name, df)
            report.sheet_profiles.append(profile)
    else:
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
    # Use raw pandas DataFrames (both sides) to avoid format mismatches
    comparisons: list[DataComparisonReport] = []

    if source_dataframes:
        # Prefer generated_dataframes (raw pandas) over CSV-parsed DataFrames
        # Exclude hidden sheets — they are not visible to the user
        _hidden = hidden_sheets or set()
        gen_dfs: dict[str, pd.DataFrame] = {}
        if generated_dataframes:
            gen_dfs = {name: df for name, df in generated_dataframes.items()
                       if len(df) > 0 and name not in _hidden}
        else:
            for name, csv_text in generated_csv_texts.items():
                if name in _hidden:
                    continue
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
                # Column overlap is more important than sheet name similarity
                score = name_sim * 0.3 + (col_overlap / col_total) * 0.7
                if score > best_score:
                    best_score = score
                    best_name = gen_name
            if best_name and best_score >= 0.2:
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
        if dc.empty_rows_dropped > 0:
            lines.append(f"- Empty rows removed before comparison: {dc.empty_rows_dropped}")

        # Matching method info
        if dc.match_method == "key_based":
            lines.append(f"- **Matching method**: Key-based (joined on column `{dc.key_column}`). "
                         f"Row order differences do NOT affect results.")
        else:
            lines.append(f"- **Matching method**: Positional (row-by-row). "
                         f"⚠ If rows are sorted differently, this will report false differences.")

        # Add context for row count differences
        if dc.source_rows > 0 and dc.generated_rows > 0:
            ratio = dc.generated_rows / dc.source_rows
            if ratio < 0.3:
                lines.append(
                    f"- **Note**: Generated file has significantly fewer rows than source "
                    f"({dc.generated_rows} vs {dc.source_rows}, {ratio:.0%}). "
                    f"This likely indicates the output is a filtered subset, not a copy of all data. "
                    f"Row-by-row match rate is NOT meaningful in this case."
                )
            elif ratio > 3:
                lines.append(
                    f"- **Note**: Generated file has significantly more rows than source "
                    f"({dc.generated_rows} vs {dc.source_rows}). "
                    f"The output likely includes calculated/derived rows beyond the source data."
                )

        lines.append(f"- Comparable rows: {dc.total_comparable_rows}, Matched: {dc.matched_rows} ({dc.match_percentage:.1f}%)")

        # Low match rate warning
        if dc.match_percentage < 50 and dc.total_comparable_rows > 0:
            if dc.match_method == "positional":
                lines.append(
                    f"- ⚠ **Low match rate ({dc.match_percentage:.1f}%) with positional matching.** "
                    f"This may be caused by different row ordering rather than actual data errors. "
                    f"Interpret with caution."
                )
            else:
                lines.append(
                    f"- ⚠ **Low match rate ({dc.match_percentage:.1f}%).** "
                    f"Many matched rows have value differences."
                )

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
    return _profile_dataframe(name, df)


def _profile_dataframe(name: str, df: pd.DataFrame) -> SheetProfile:
    """Profile a sheet from a pandas DataFrame."""

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
    """Compare source and generated DataFrames.

    Steps:
    1. Drop all-empty rows from both sides
    2. Match columns by name/value similarity
    3. Try key-based row matching (find a unique identifier column)
    4. Fall back to positional matching if no key column found
    """
    # Step 0: Drop all-empty rows from both sides
    src_clean, src_dropped = _drop_empty_rows(source_df)
    gen_clean, gen_dropped = _drop_empty_rows(gen_df)

    report = DataComparisonReport(
        source_rows=len(src_clean),
        generated_rows=len(gen_clean),
        empty_rows_dropped=src_dropped + gen_dropped,
    )

    # Step 1: Match columns
    col_mappings = _match_columns(src_clean, gen_clean)
    report.column_mappings = col_mappings

    if not col_mappings:
        report.summary = "No column mappings could be established between source and generated data."
        return report

    # Step 2: Filter to high-confidence column mappings
    mapped_src_cols = [cm.source_col for cm in col_mappings if cm.confidence >= 0.5]
    mapped_gen_cols = [cm.generated_col for cm in col_mappings if cm.confidence >= 0.5]

    if not mapped_src_cols:
        report.summary = "Column mappings too low confidence for row comparison."
        return report

    # Step 3: Try key-based matching
    key_col = _find_key_column(src_clean, gen_clean, col_mappings)
    if key_col:
        report.match_method = "key_based"
        report.key_column = key_col[0]  # source key column name
        _compare_by_key(src_clean, gen_clean, key_col, mapped_src_cols, mapped_gen_cols, report)
    else:
        report.match_method = "positional"
        _compare_positional(src_clean, gen_clean, mapped_src_cols, mapped_gen_cols, report)

    return report


def _drop_empty_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Drop rows where all values are NaN or empty string. Returns (clean_df, dropped_count)."""
    def _is_cell_empty(val) -> bool:
        if pd.isna(val):
            return True
        return str(val).strip() == ""

    mask = df.apply(lambda row: all(_is_cell_empty(v) for v in row), axis=1)
    dropped = int(mask.sum())
    return df[~mask].reset_index(drop=True), dropped


def _find_key_column(
    source_df: pd.DataFrame,
    gen_df: pd.DataFrame,
    col_mappings: list[ColumnMapping],
) -> tuple[str, str] | None:
    """Find a suitable key column for row matching.

    Returns (source_col, generated_col) if found, None otherwise.
    Criteria: uniqueness > 90% on both sides.
    For overlap: check that most generated values exist in source (directional),
    not symmetric overlap — handles filtered subsets where gen << source.
    """
    for cm in col_mappings:
        if cm.confidence < 0.5:
            continue

        src_col, gen_col = cm.source_col, cm.generated_col

        src_series = source_df[src_col].dropna()
        gen_series = gen_df[gen_col].dropna()

        if len(src_series) == 0 or len(gen_series) == 0:
            continue

        src_uniqueness = src_series.nunique() / len(src_series)
        gen_uniqueness = gen_series.nunique() / len(gen_series)

        if src_uniqueness < 0.9 or gen_uniqueness < 0.9:
            continue

        # Directional overlap: what fraction of generated values appear in source?
        src_vals = set(src_series.astype(str))
        gen_vals = set(gen_series.astype(str))
        if len(gen_vals) == 0:
            continue
        gen_in_src = len(gen_vals & src_vals) / len(gen_vals)

        if gen_in_src >= 0.5:
            return (src_col, gen_col)

    return None


def _compare_by_key(
    source_df: pd.DataFrame,
    gen_df: pd.DataFrame,
    key_col: tuple[str, str],
    mapped_src_cols: list[str],
    mapped_gen_cols: list[str],
    report: DataComparisonReport,
) -> None:
    """Compare DataFrames by joining on a key column."""
    src_key, gen_key = key_col

    # Build lookup: key_value → row index in generated
    gen_lookup: dict[str, int] = {}
    for idx in range(len(gen_df)):
        key_val = str(gen_df.iloc[idx][gen_key])
        if key_val not in gen_lookup:
            gen_lookup[key_val] = idx

    compare_rows = 0
    matched = 0
    diffs: list[RowDiff] = []

    for src_idx in range(len(source_df)):
        key_val = str(source_df.iloc[src_idx][src_key])
        gen_idx = gen_lookup.get(key_val)
        if gen_idx is None:
            continue  # key not found in generated — skip (could be filtered out)

        compare_rows += 1
        row_match = True
        for src_col, gen_col in zip(mapped_src_cols, mapped_gen_cols):
            src_val = source_df.iloc[src_idx].get(src_col)
            gen_val = gen_df.iloc[gen_idx].get(gen_col)

            if not _values_match(src_val, gen_val):
                row_match = False
                if len(diffs) < 50:
                    diffs.append(RowDiff(
                        row_index=src_idx + 2,  # 1-indexed + header
                        column=src_col,
                        source_value=str(src_val)[:100],
                        generated_value=str(gen_val)[:100],
                    ))

        if row_match:
            matched += 1

    report.total_comparable_rows = compare_rows
    report.matched_rows = matched
    report.match_percentage = (matched / compare_rows * 100) if compare_rows > 0 else 0
    report.row_diffs = diffs

    diff_count = compare_rows - matched
    unmatched_keys = len(source_df) - compare_rows
    report.summary = (
        f"Key-based matching on `{src_key}`: {matched}/{compare_rows} rows match "
        f"({report.match_percentage:.1f}%). "
        f"{diff_count} rows have differences across {len(mapped_src_cols)} compared columns."
    )
    if unmatched_keys > 0:
        report.summary += f" {unmatched_keys} source rows had no matching key in generated data."

    # Update value_match_pct on column mappings
    _update_column_match_pct_keyed(
        source_df, gen_df, src_key, gen_key, gen_lookup,
        report.column_mappings, mapped_src_cols, mapped_gen_cols,
    )


def _compare_positional(
    source_df: pd.DataFrame,
    gen_df: pd.DataFrame,
    mapped_src_cols: list[str],
    mapped_gen_cols: list[str],
    report: DataComparisonReport,
) -> None:
    """Compare DataFrames row-by-row by position (fallback when no key column found)."""
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
    for cm in report.column_mappings:
        if cm.source_col in mapped_src_cols:
            idx = mapped_src_cols.index(cm.source_col)
            gen_col = mapped_gen_cols[idx]
            match_count = sum(
                1 for i in range(compare_rows)
                if _values_match(source_df.iloc[i].get(cm.source_col), gen_df.iloc[i].get(gen_col))
            )
            cm.value_match_pct = match_count / compare_rows * 100 if compare_rows > 0 else 0


def _update_column_match_pct_keyed(
    source_df: pd.DataFrame,
    gen_df: pd.DataFrame,
    src_key: str,
    gen_key: str,
    gen_lookup: dict[str, int],
    col_mappings: list[ColumnMapping],
    mapped_src_cols: list[str],
    mapped_gen_cols: list[str],
) -> None:
    """Update value_match_pct on column mappings for key-based comparison."""
    for cm in col_mappings:
        if cm.source_col not in mapped_src_cols:
            continue
        idx = mapped_src_cols.index(cm.source_col)
        gen_col = mapped_gen_cols[idx]
        match_count = 0
        total = 0
        for src_idx in range(len(source_df)):
            key_val = str(source_df.iloc[src_idx][src_key])
            gen_idx = gen_lookup.get(key_val)
            if gen_idx is None:
                continue
            total += 1
            if _values_match(source_df.iloc[src_idx].get(cm.source_col), gen_df.iloc[gen_idx].get(gen_col)):
                match_count += 1
        cm.value_match_pct = match_count / total * 100 if total > 0 else 0


def _match_columns(source_df: pd.DataFrame, gen_df: pd.DataFrame) -> list[ColumnMapping]:
    """Match source columns to generated columns using name and value similarity.

    Two-pass approach: exact name matches first (to avoid greedy mismatches),
    then fuzzy/value-based matches for remaining columns.
    """
    mappings: list[ColumnMapping] = []
    used_gen_cols: set[str] = set()
    matched_src_cols: set[str] = set()

    # Pass 1: Exact name matches (case-insensitive)
    gen_col_lower_map: dict[str, str] = {str(c).lower(): str(c) for c in gen_df.columns}
    for src_col in source_df.columns:
        src_lower = str(src_col).lower()
        if src_lower in gen_col_lower_map:
            gen_col = gen_col_lower_map[src_lower]
            if gen_col in used_gen_cols:
                continue
            name_sim = SequenceMatcher(None, src_lower, str(gen_col).lower()).ratio()
            mappings.append(ColumnMapping(
                source_col=str(src_col),
                generated_col=gen_col,
                match_type="exact_name",
                confidence=max(name_sim, 0.95),
            ))
            used_gen_cols.add(gen_col)
            matched_src_cols.add(str(src_col))

    # Pass 2: Fuzzy/value-based matches for remaining columns
    for src_col in source_df.columns:
        if str(src_col) in matched_src_cols:
            continue

        best_match: ColumnMapping | None = None
        best_score = 0.0

        for gen_col in gen_df.columns:
            if str(gen_col) in used_gen_cols:
                continue

            name_sim = SequenceMatcher(None, str(src_col).lower(), str(gen_col).lower()).ratio()

            src_vals = set(source_df[src_col].dropna().astype(str).head(100))
            gen_vals = set(gen_df[gen_col].dropna().astype(str).head(100))
            val_overlap = len(src_vals & gen_vals) / max(len(src_vals | gen_vals), 1)

            score = name_sim * 0.6 + val_overlap * 0.4

            if score > best_score:
                best_score = score
                match_type = "fuzzy_name" if name_sim > 0.5 else "value_overlap"
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

    # String comparison (strip whitespace)
    str_a, str_b = str(a).strip(), str(b).strip()
    if str_a == str_b:
        return True

    # Strip thousand separators and currency symbols before numeric comparison
    clean_a = str_a.replace(",", "").replace("$", "").replace("€", "").replace("¥", "")
    clean_b = str_b.replace(",", "").replace("$", "").replace("€", "").replace("¥", "")
    if clean_a == clean_b:
        return True

    # Numeric comparison with tolerance
    try:
        num_a, num_b = float(clean_a), float(clean_b)
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
