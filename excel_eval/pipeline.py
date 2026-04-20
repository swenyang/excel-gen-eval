"""Pipeline orchestration — wires Stage 1 → Stage 2 → Stage 3."""

from __future__ import annotations

import asyncio
import logging
from datetime import datetime, timezone
from pathlib import Path

from .config import load_case_config, discover_cases
from .evaluators import ALL_EVALUATORS, ScenarioDetector
from .evaluators.base import BaseEvaluator
from .llm import create_llm_client
from .llm import create_llm_client
from .models import (
    CaseConfig,
    CostSummary,
    DimensionName,
    DimensionResult,
    EvalConfig,
    EvalResult,
    EvalStatus,
    EvalSummary,
    GlobalConfig,
    PreparedData,
    Scenario,
    ScenarioResult,
    get_blended_weights,
    get_weights,
)
from .parsers.excel_parser import parse_excel
from .parsers.input_loader import load_all_inputs
from .parsers.screenshot import generate_screenshots

logger = logging.getLogger(__name__)


DATA_CONTENT_DIMS = {
    DimensionName.DATA_ACCURACY,
    DimensionName.COMPLETENESS,
    DimensionName.FORMULA_LOGIC,
    DimensionName.RELEVANCE,
}
STRUCTURE_USABILITY_DIMS = {
    DimensionName.SHEET_ORGANIZATION,
    DimensionName.TABLE_STRUCTURE,
    DimensionName.CHART_APPROPRIATENESS,
    DimensionName.PROFESSIONAL_FORMATTING,
}


class Pipeline:
    """End-to-end evaluation pipeline."""

    def __init__(self, config: GlobalConfig) -> None:
        self.config = config
        self.llm_client = create_llm_client(config.llm)

    # ── Public API ─────────────────────────────────────────────────────

    async def evaluate(
        self, case_dir: str | Path, num_runs: int = 1,
        output_dir: str | Path | None = None,
    ) -> EvalResult:
        """Evaluate a single test case.

        Args:
            case_dir: Path to the test case directory.
            num_runs: Number of evaluation runs. If >1, takes the median
                      score per dimension for stability.
            output_dir: If provided, saves Stage 1 debug data to output_dir/debug/
        """
        case_dir = Path(case_dir)
        case_config = load_case_config(case_dir)
        logger.info("Evaluating case: %s", case_config.id)

        import time as _time
        t_start = _time.perf_counter()

        # Stage 1: Data Preparation (once)
        t_s1 = _time.perf_counter()
        prepared = self._stage1_prepare(case_dir, case_config)
        stage1_ms = int((_time.perf_counter() - t_s1) * 1000)

        # Save debug data if output_dir provided
        if output_dir:
            self._save_debug_data(Path(output_dir), case_config, prepared)

        # Stage 2.0: Scenario Detection (once)
        scenario = await self._detect_scenario(prepared, case_config)

        if num_runs <= 1:
            # Single run
            dim_results = await self._stage2_evaluate(
                prepared, scenario.detected, case_config
            )
        else:
            # Multi-run: evaluate N times, take median per dimension
            logger.info("Running %d evaluation passes for stability", num_runs)
            all_runs: list[dict[str, DimensionResult]] = []
            for run_i in range(num_runs):
                logger.info("  Pass %d/%d", run_i + 1, num_runs)
                run_results = await self._stage2_evaluate(
                    prepared, scenario.detected, case_config
                )
                all_runs.append(run_results)
            dim_results = self._merge_runs_median(all_runs)

        # Stage 3: Aggregation
        result = self._stage3_aggregate(
            case_config, prepared, scenario, dim_results
        )

        total_ms = int((_time.perf_counter() - t_start) * 1000)
        result.metadata["performance"] = {
            "stage1_ms": stage1_ms,
            "stage2_llm_ms": result.cost.total_latency_ms,
            "total_wall_ms": total_ms,
            "screenshots_count": len(prepared.screenshots),
        }

        logger.info(
            "Case %s complete: overall=%.2f (%s)",
            case_config.id,
            result.summary.overall_weighted_avg or 0,
            scenario.detected.value,
        )
        return result

    async def evaluate_batch(
        self, root_dir: str | Path, num_runs: int = 1,
        parallel: int = 1, output_dir: str | Path | None = None,
    ) -> list[EvalResult]:
        """Evaluate all test cases under root_dir.

        Results are saved incrementally — each completed case is immediately
        written to ``output_dir/results/{case_id}.json``. On restart, already
        completed cases are loaded from disk and skipped, enabling resumable
        batch runs.

        Args:
            root_dir: Directory containing test case subdirectories.
            num_runs: Number of runs per case (>1 uses median).
            parallel: Number of cases to evaluate concurrently.
            output_dir: If provided, saves debug data and incremental results.
        """
        cases = discover_cases(root_dir)
        if not cases:
            logger.warning("No test cases found under %s", root_dir)
            return []

        logger.info("Found %d test case(s), parallel=%d", len(cases), parallel)

        def _case_output_dir(case_dir: Path) -> Path | None:
            if not output_dir:
                return None
            return Path(output_dir) / "debug" / case_dir.name

        # Incremental results directory
        results_dir = Path(output_dir) / "results" if output_dir else None
        if results_dir:
            results_dir.mkdir(parents=True, exist_ok=True)

        # Load already-completed results (for resume support)
        completed: dict[str, EvalResult] = {}
        if results_dir:
            import json as _json
            for result_file in results_dir.glob("*.json"):
                try:
                    with open(result_file, "r", encoding="utf-8") as f:
                        data = _json.load(f)
                    er = EvalResult(**data)
                    completed[er.case_id] = er
                except Exception as exc:
                    logger.warning("Could not load cached result %s: %s", result_file, exc)
            if completed:
                logger.info("Resuming: %d case(s) already completed, %d remaining",
                            len(completed), len(cases) - len(completed))

        def _save_result(result: EvalResult) -> None:
            """Save a single case result to disk immediately."""
            if not results_dir:
                return
            import json as _json
            result_path = results_dir / f"{result.case_id}.json"
            with open(result_path, "w", encoding="utf-8") as f:
                _json.dump(
                    _json.loads(result.model_dump_json()),
                    f, indent=2, ensure_ascii=False, default=str,
                )

        async def _eval_one(case_dir: Path) -> EvalResult | None:
            """Evaluate a single case with timeout, save result immediately."""
            # Skip already-completed cases
            case_id = case_dir.name
            if case_id in completed:
                logger.info("Skipping %s (already completed)", case_id)
                return completed[case_id]

            try:
                result = await asyncio.wait_for(
                    self.evaluate(
                        case_dir, num_runs=num_runs,
                        output_dir=_case_output_dir(case_dir),
                    ),
                    timeout=600,  # 10 min max per case
                )
                _save_result(result)
                return result
            except asyncio.TimeoutError:
                logger.error("Case %s timed out after 600s", case_dir.name)
                return None
            except Exception as exc:
                logger.exception("Failed to evaluate case %s: %s", case_dir, exc)
                return None

        if parallel <= 1:
            results: list[EvalResult] = []
            for case_dir in cases:
                result = await _eval_one(case_dir)
                if result is not None:
                    results.append(result)
            return results

        semaphore = asyncio.Semaphore(parallel)

        async def _run_one(case_dir: Path) -> EvalResult | None:
            async with semaphore:
                return await _eval_one(case_dir)

        tasks = [_run_one(case_dir) for case_dir in cases]
        raw_results = await asyncio.gather(*tasks)
        return [r for r in raw_results if r is not None]

    # ── Stage 1: Data Preparation ──────────────────────────────────────

    def _stage1_prepare(
        self, case_dir: Path, case_config: CaseConfig
    ) -> PreparedData:
        """Parse Excel + load grounding data + generate screenshots."""
        logger.info("Stage 1: Preparing data for case %s", case_config.id)

        # Find the Excel file
        excel_path = self._resolve_excel_path(case_dir, case_config)

        # Load grounding data
        grounding = ""
        if case_config.input_files:
            input_configs = [f.model_dump() for f in case_config.input_files]
            grounding = load_all_inputs(input_configs, case_dir)

        # Parse Excel
        prepared = parse_excel(
            excel_path,
            grounding_data=grounding,
            user_prompt=case_config.prompt,
        )

        # Generate screenshots (required by default)
        screenshots = generate_screenshots(excel_path, required=self.config.evaluation.screenshot_enabled)
        prepared.screenshots = screenshots
        if screenshots:
            logger.info("Generated %d screenshot(s) for sheets: %s",
                         len(screenshots), ", ".join(screenshots.keys()))
        else:
            logger.warning("No screenshots generated")

        # Run code-level data scanner
        from .parsers.data_scanner import scan_generated_excel, format_scan_report
        import pandas as pd

        csv_texts = {s.name: s.csv_text for s in prepared.sheets}

        # Load generated DataFrames via pandas (raw values, for scanner comparison)
        generated_dfs: dict[str, pd.DataFrame] = {}
        try:
            gen_xls = pd.ExcelFile(excel_path)
            for sheet_name in gen_xls.sheet_names:
                generated_dfs[sheet_name] = pd.read_excel(gen_xls, sheet_name=sheet_name)
        except Exception as e:
            logger.warning("Could not load generated Excel for scanning: %s", e)

        # Load source DataFrames from input files for precise comparison
        source_dfs: dict[str, pd.DataFrame] = {}
        if case_config.input_files:
            from .parsers.input_loader import extract_dataframes
            for fc in case_config.input_files:
                input_path = case_dir / fc.path
                if input_path.exists():
                    try:
                        dfs = extract_dataframes(input_path)
                        if dfs:
                            source_dfs.update(dfs)
                            logger.info("Loaded %d source table(s) from %s for comparison",
                                         len(dfs), fc.path)
                    except Exception as e:
                        logger.warning("Could not load source %s: %s", fc.path, e)

        hidden_sheet_names = {s.name for s in prepared.sheets if s.hidden}

        scan = scan_generated_excel(
            csv_texts,
            source_text=grounding,
            source_dataframes=source_dfs if source_dfs else None,
            generated_dataframes=generated_dfs if generated_dfs else None,
            formulas=prepared.formulas,
            hidden_sheets=hidden_sheet_names if hidden_sheet_names else None,
        )
        prepared.scan_report_text = format_scan_report(scan)
        logger.info("Data scan complete: %d sheet profiles, %d formula errors, comparison=%s",
                     len(scan.sheet_profiles), scan.formula_error_count,
                     "yes" if scan.data_comparison and scan.data_comparison.total_comparable_rows > 0 else "no")

        return prepared

    # ── Stage 2: Evaluation ────────────────────────────────────────────

    async def _detect_scenario(
        self, data: PreparedData, case_config: CaseConfig
    ) -> ScenarioResult:
        """Detect scenario via LLM or use manual override."""
        if case_config.scenario:
            try:
                manual = Scenario(case_config.scenario)
                logger.info("Using manual scenario override: %s", manual.value)
                return ScenarioResult(
                    detected=manual, confidence=1.0, reasoning="Manual override"
                )
            except ValueError:
                logger.warning(
                    "Invalid scenario override '%s', auto-detecting",
                    case_config.scenario,
                )

        detector = ScenarioDetector(self.llm_client)
        return await detector.detect(data)

    async def _stage2_evaluate(
        self,
        data: PreparedData,
        scenario: Scenario,
        case_config: CaseConfig,
    ) -> dict[str, DimensionResult]:
        """Run all dimension evaluations in parallel."""
        logger.info("Stage 2: Evaluating %d dimensions", len(ALL_EVALUATORS))

        skip_set = set(case_config.skip_dimensions)
        semaphore = asyncio.Semaphore(self.config.evaluation.max_concurrent_calls)

        async def _run_one(evaluator: BaseEvaluator) -> DimensionResult:
            dim_name = evaluator.dimension.value
            if dim_name in skip_set:
                return DimensionResult(
                    dimension=evaluator.dimension,
                    status=EvalStatus.SKIPPED,
                    feedback="Skipped by configuration",
                )
            # Error on screenshot-dependent dimensions when no screenshots available
            screenshot_required_dims = {"professional_formatting", "chart_appropriateness"}
            if dim_name in screenshot_required_dims and not data.screenshots:
                return DimensionResult(
                    dimension=evaluator.dimension,
                    score=None,
                    status=EvalStatus.ERROR,
                    feedback="N/A — no screenshots available (LibreOffice/poppler conversion failed). "
                             "This dimension requires visual inspection and cannot be evaluated from CSV data alone.",
                    error_message="Screenshot generation failed — install LibreOffice and poppler (pdf2image) "
                                  "to enable visual evaluation dimensions.",
                )
            async with semaphore:
                logger.info("  Evaluating: %s", dim_name)
                return await evaluator.evaluate(data, scenario)

        language = self.config.evaluation.language
        evaluators = [cls(self.llm_client, language=language) for cls in ALL_EVALUATORS]
        tasks = [_run_one(e) for e in evaluators]
        results = await asyncio.gather(*tasks)

        return {r.dimension.value: r for r in results}

    # ── Stage 3: Aggregation ───────────────────────────────────────────

    def _stage3_aggregate(
        self,
        case_config: CaseConfig,
        prepared: PreparedData,
        scenario: ScenarioResult,
        dim_results: dict[str, DimensionResult],
    ) -> EvalResult:
        """Aggregate dimension results into final EvalResult."""
        # Cross-dimension consistency adjustment
        dim_results = self._apply_consistency_checks(dim_results)

        weights = get_blended_weights(scenario)

        # Compute block averages
        data_content_scores: list[tuple[float, float]] = []
        structure_scores: list[tuple[float, float]] = []
        all_weighted: list[tuple[float, float]] = []
        na_dims: list[str] = []
        skipped_dims: list[str] = []

        for dim_name, result in dim_results.items():
            if result.status == EvalStatus.SKIPPED:
                skipped_dims.append(dim_name)
                continue
            if result.status == EvalStatus.ERROR or result.score is None:
                na_dims.append(dim_name)
                continue

            w = weights.get(dim_name, 1.0)
            weighted_score = (result.score, w)
            all_weighted.append(weighted_score)

            dim_enum = DimensionName(dim_name)
            if dim_enum in DATA_CONTENT_DIMS:
                data_content_scores.append(weighted_score)
            elif dim_enum in STRUCTURE_USABILITY_DIMS:
                structure_scores.append(weighted_score)

        summary = EvalSummary(
            data_content_avg=_weighted_avg(data_content_scores),
            structure_usability_avg=_weighted_avg(structure_scores),
            overall_weighted_avg=_weighted_avg(all_weighted),
            dimensions_evaluated=len(all_weighted),
            na_dimensions=na_dims,
            skipped_dimensions=skipped_dims,
            weights_applied={k: weights.get(k, 1.0) for k in dim_results},
        )

        # Cost summary
        cost = CostSummary(
            total_input_tokens=sum(r.input_tokens for r in dim_results.values()),
            total_output_tokens=sum(r.output_tokens for r in dim_results.values()),
            total_latency_ms=max(
                (r.latency_ms for r in dim_results.values()), default=0
            ),
            total_cost_estimate=sum(r.cost_estimate for r in dim_results.values()),
            per_dimension={
                name: {
                    "input_tokens": r.input_tokens,
                    "output_tokens": r.output_tokens,
                    "latency_ms": r.latency_ms,
                    "cost_estimate": r.cost_estimate,
                }
                for name, r in dim_results.items()
            },
        )

        excel_file = ""
        if case_config.output_files:
            excel_file = case_config.output_files[0].path

        return EvalResult(
            case_id=case_config.id,
            timestamp=datetime.now(timezone.utc),
            excel_file=excel_file,
            prompt=case_config.prompt,
            input_files=[f.model_dump() for f in case_config.input_files],
            output_files=[f.model_dump() for f in case_config.output_files],
            metadata=case_config.metadata,
            scenario=scenario,
            dimensions=dim_results,
            summary=summary,
            cost=cost,
            llm_config={
                "provider": self.config.llm.provider,
                "model": self.config.llm.model,
                "temperature": self.config.llm.temperature,
            },
        )

    # ── Helpers ────────────────────────────────────────────────────────

    @staticmethod
    def _merge_runs_median(
        all_runs: list[dict[str, DimensionResult]],
    ) -> dict[str, DimensionResult]:
        """Merge multiple evaluation runs by taking the median score per dimension.

        Uses the result whose score equals the median as the representative
        (preserving its feedback and evidence).
        """
        import statistics

        dim_names = all_runs[0].keys()
        merged: dict[str, DimensionResult] = {}

        for dim_name in dim_names:
            run_results = [run[dim_name] for run in all_runs]
            scores = [r.score for r in run_results if r.score is not None]

            if not scores:
                # All runs returned N/A — use first
                merged[dim_name] = run_results[0]
                continue

            median_score = int(statistics.median(scores))

            # Find the run closest to median to use as representative
            best = min(
                (r for r in run_results if r.score is not None),
                key=lambda r: abs(r.score - median_score),
            )

            if best.score != median_score:
                best.score = median_score
                best.feedback += (
                    f" [Median of {len(scores)} runs: scores were "
                    f"{sorted(scores)}, median={median_score}]"
                )

            merged[dim_name] = best

        return merged

    @staticmethod
    def _apply_consistency_checks(
        dim_results: dict[str, DimensionResult],
    ) -> dict[str, DimensionResult]:
        """Apply cross-dimension consistency adjustments.

        Rules:
        - If Data Accuracy = 0 (data is fully fabricated/hallucinated), cap
          Completeness at max 2 because "complete but fabricated" is misleading.
        - If Formula & Logic = 0 (critical formula errors), cap Data Accuracy
          at max 2 because formula errors cascade into data errors.
        """
        da = dim_results.get(DimensionName.DATA_ACCURACY)
        comp = dim_results.get(DimensionName.COMPLETENESS)
        fl = dim_results.get(DimensionName.FORMULA_LOGIC)

        if da and da.score is not None and da.score == 0:
            # Data is fully fabricated — Completeness shouldn't be high
            if comp and comp.score is not None and comp.score > 2:
                old = comp.score
                comp.score = 2
                comp.feedback += (
                    f" [Consistency adjustment: Completeness capped from {old} to "
                    f"{comp.score} because Data Accuracy is 0 — fabricated content "
                    f"has limited value.]"
                )

        if fl and fl.score is not None and fl.score == 0:
            # Critical formula errors — Data Accuracy may be overstated
            if da and da.score is not None and da.score > 2:
                old = da.score
                da.score = 2
                da.feedback += (
                    f" [Consistency adjustment: Data Accuracy capped from {old} to "
                    f"2 because Formula & Logic is 0 — critical formula errors "
                    f"likely compromise data correctness.]"
                )

        return dim_results

    @staticmethod
    def _save_debug_data(
        output_dir: Path, case_config: CaseConfig, prepared: PreparedData,
    ) -> None:
        """Save Stage 1 intermediate data for debugging."""
        import json as _json
        import re as _re

        debug_dir = output_dir
        debug_dir.mkdir(parents=True, exist_ok=True)

        def _safe_filename(name: str) -> str:
            """Remove characters illegal in Windows filenames."""
            return _re.sub(r'[<>:"/\\|?*]', '_', name)

        # Save CSV per sheet
        for sheet in prepared.sheets:
            csv_path = debug_dir / f"sheet_{_safe_filename(sheet.name)}.csv"
            csv_path.write_text(sheet.csv_text, encoding="utf-8")

        # Save scan report
        if prepared.scan_report_text:
            (debug_dir / "scan_report.txt").write_text(
                prepared.scan_report_text, encoding="utf-8"
            )

        # Save screenshots
        for name, img_bytes in prepared.screenshots.items():
            (debug_dir / f"screenshot_{_safe_filename(name)}.png").write_bytes(img_bytes)

        # Save formula list
        if prepared.formulas:
            formulas = [
                {"sheet": f.sheet, "cell": f.cell, "formula": f.formula,
                 "value": str(f.computed_value), "error": f.has_error}
                for f in prepared.formulas[:500]
            ]
            (debug_dir / "formulas.json").write_text(
                _json.dumps(formulas, indent=2, ensure_ascii=False), encoding="utf-8"
            )

        # Save metadata summary
        meta = {
            "case_id": case_config.id,
            "sheets": [{"name": s.name, "rows": s.row_count, "cols": s.col_count,
                         "truncated": s.truncated} for s in prepared.sheets],
            "total_formulas": len(prepared.formulas),
            "formula_errors": sum(1 for f in prepared.formulas if f.has_error),
            "charts": [{"sheet": c.sheet, "type": c.chart_type, "title": c.title}
                        for c in prepared.charts],
            "cross_sheet_refs": prepared.cross_sheet_refs[:20],
            "screenshots_count": len(prepared.screenshots),
            "formatting": {
                "fonts": prepared.formatting.fonts_used,
                "colors": prepared.formatting.color_palette[:10],
                "conditional_formatting": prepared.formatting.has_conditional_formatting,
                "frozen_panes": prepared.formatting.frozen_panes,
                "merged_cells": len(prepared.formatting.merged_cell_ranges),
            },
        }
        (debug_dir / "stage1_meta.json").write_text(
            _json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8"
        )

        logger.info("Debug data saved to %s", debug_dir)

    @staticmethod
    def _resolve_excel_path(case_dir: Path, case_config: CaseConfig) -> Path:
        """Find the Excel file from config or by scanning output dir."""
        if case_config.output_files:
            path = case_dir / case_config.output_files[0].path
            if path.exists():
                return path

        # Fallback: scan for .xlsx files
        for xlsx in case_dir.rglob("*.xlsx"):
            return xlsx

        raise FileNotFoundError(f"No Excel file found in {case_dir}")


def _weighted_avg(scores: list[tuple[float, float]]) -> float | None:
    """Compute weighted average from (score, weight) pairs."""
    if not scores:
        return None
    total_weight = sum(w for _, w in scores)
    if total_weight == 0:
        return None
    return sum(s * w for s, w in scores) / total_weight
