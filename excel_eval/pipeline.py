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
    ) -> EvalResult:
        """Evaluate a single test case.

        Args:
            case_dir: Path to the test case directory.
            num_runs: Number of evaluation runs. If >1, takes the median
                      score per dimension for stability.
        """
        case_dir = Path(case_dir)
        case_config = load_case_config(case_dir)
        logger.info("Evaluating case: %s", case_config.id)

        # Stage 1: Data Preparation (once)
        prepared = self._stage1_prepare(case_dir, case_config)

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

        logger.info(
            "Case %s complete: overall=%.2f (%s)",
            case_config.id,
            result.summary.overall_weighted_avg or 0,
            scenario.detected.value,
        )
        return result

    async def evaluate_batch(
        self, root_dir: str | Path, num_runs: int = 1,
    ) -> list[EvalResult]:
        """Evaluate all test cases under root_dir."""
        cases = discover_cases(root_dir)
        if not cases:
            logger.warning("No test cases found under %s", root_dir)
            return []

        logger.info("Found %d test case(s)", len(cases))
        results: list[EvalResult] = []

        for case_dir in cases:
            try:
                result = await self.evaluate(case_dir, num_runs=num_runs)
                results.append(result)
            except Exception as exc:
                logger.exception("Failed to evaluate case %s: %s", case_dir, exc)

        return results

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

        # Load source DataFrames from input Excel files for precise comparison
        source_dfs: dict[str, pd.DataFrame] = {}
        if case_config.input_files:
            for fc in case_config.input_files:
                input_path = case_dir / fc.path
                if input_path.suffix.lower() in (".xlsx", ".xls") and input_path.exists():
                    try:
                        xls = pd.ExcelFile(input_path)
                        for sheet_name in xls.sheet_names:
                            source_dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
                        logger.info("Loaded %d source sheet(s) from %s for comparison",
                                     len(xls.sheet_names), fc.path)
                    except Exception as e:
                        logger.warning("Could not load source Excel %s: %s", fc.path, e)

        scan = scan_generated_excel(
            csv_texts,
            source_text=grounding,
            source_dataframes=source_dfs if source_dfs else None,
            generated_dataframes=generated_dfs if generated_dfs else None,
            formulas=prepared.formulas,
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
            async with semaphore:
                logger.info("  Evaluating: %s", dim_name)
                return await evaluator.evaluate(data, scenario)

        evaluators = [cls(self.llm_client) for cls in ALL_EVALUATORS]
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
        - If Data Accuracy <= 1 (data is fundamentally wrong), cap Completeness
          at DA+1 because "complete but wrong" is misleading.
        - If Formula & Logic = 0 (critical formula errors), cap Data Accuracy
          at max 2 because formula errors cascade into data errors.
        """
        da = dim_results.get(DimensionName.DATA_ACCURACY)
        comp = dim_results.get(DimensionName.COMPLETENESS)
        fl = dim_results.get(DimensionName.FORMULA_LOGIC)

        if da and da.score is not None and da.score <= 1:
            # Data is fundamentally wrong — Completeness shouldn't be high
            if comp and comp.score is not None and comp.score > da.score + 1:
                old = comp.score
                comp.score = da.score + 1
                comp.feedback += (
                    f" [Consistency adjustment: Completeness capped from {old} to "
                    f"{comp.score} because Data Accuracy is {da.score} — content "
                    f"that is present but incorrect has limited value.]"
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
