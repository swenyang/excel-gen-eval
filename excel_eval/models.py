"""Data models for the Excel evaluation pipeline."""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


# ── Enums ──────────────────────────────────────────────────────────────────


class Scenario(str, Enum):
    REPORTING_ANALYSIS = "reporting_analysis"
    DATA_PROCESSING = "data_processing"
    TEMPLATE_FORM = "template_form"
    PLANNING_TRACKING = "planning_tracking"
    FINANCIAL_MODELING = "financial_modeling"
    COMPARISON_EVALUATION = "comparison_evaluation"
    GENERAL = "general"


class DimensionName(str, Enum):
    DATA_ACCURACY = "data_accuracy"
    COMPLETENESS = "completeness"
    FORMULA_LOGIC = "formula_logic"
    RELEVANCE = "relevance"
    SHEET_ORGANIZATION = "sheet_organization"
    TABLE_STRUCTURE = "table_structure"
    CHART_APPROPRIATENESS = "chart_appropriateness"
    PROFESSIONAL_FORMATTING = "professional_formatting"


class EvalStatus(str, Enum):
    SUCCESS = "success"
    ERROR = "error"
    NA = "na"
    SKIPPED = "skipped"


# ── Stage 1: Prepared Data ─────────────────────────────────────────────────


class SheetData(BaseModel):
    """CSV export and metadata for a single sheet."""
    name: str
    csv_text: str
    row_count: int
    col_count: int
    truncated: bool = False
    hidden: bool = False  # True if sheet is hidden or very hidden in Excel


class FormulaInfo(BaseModel):
    """A single formula extracted from the workbook."""
    cell: str  # e.g., "B5"
    sheet: str
    formula: str  # e.g., "=SUM(B2:B4)"
    computed_value: Any = None
    has_error: bool = False


class ChartInfo(BaseModel):
    """Metadata for a single chart."""
    sheet: str
    chart_type: str  # e.g., "bar", "line", "pie"
    title: str | None = None
    data_range: str | None = None
    has_legend: bool = False
    has_axis_labels: bool = False


class FormatInfo(BaseModel):
    """Workbook-level formatting metadata."""
    fonts_used: list[str] = Field(default_factory=list)
    color_palette: list[str] = Field(default_factory=list)
    has_conditional_formatting: bool = False
    conditional_format_rules: list[str] = Field(default_factory=list)
    merged_cell_ranges: list[str] = Field(default_factory=list)
    frozen_panes: dict[str, str] = Field(default_factory=dict)  # sheet → freeze point
    has_bold_headers: bool = False  # header row uses bold
    has_borders: bool = False  # cells have border styling
    border_summary: str = ""  # e.g. "All data cells have thin borders"


class PreparedData(BaseModel):
    """All extracted data from Stage 1, consumed by Stage 2 evaluators."""
    sheets: list[SheetData] = Field(default_factory=list)
    formulas: list[FormulaInfo] = Field(default_factory=list)
    charts: list[ChartInfo] = Field(default_factory=list)
    formatting: FormatInfo = Field(default_factory=FormatInfo)
    cross_sheet_refs: list[str] = Field(default_factory=list)
    screenshots: dict[str, bytes] = Field(default_factory=dict)  # sheet_name → PNG bytes
    scan_report_text: str = ""  # Code-level scan report (factual, from data_scanner)
    grounding_data: str = ""
    user_prompt: str = ""

    @property
    def visible_sheets(self) -> list[SheetData]:
        """Return only visible (non-hidden) sheets for LLM evaluation."""
        return [s for s in self.sheets if not s.hidden]


# ── Stage 2: Evaluation Results ────────────────────────────────────────────


class DimensionResult(BaseModel):
    """Result from a single dimension evaluation."""
    dimension: DimensionName
    status: EvalStatus = EvalStatus.SUCCESS
    score: int | None = None  # 0-4, None if N/A or error
    feedback: str = ""
    evidence: list[str] = Field(default_factory=list)
    feedback_zh: str = ""  # Chinese translation
    evidence_zh: list[str] = Field(default_factory=list)
    error_message: str | None = None
    # LLM call metadata
    input_tokens: int = 0
    output_tokens: int = 0
    latency_ms: int = 0
    cost_estimate: float = 0.0


class ScenarioResult(BaseModel):
    """Result from scenario auto-detection. Supports multi-scenario blending."""
    detected: Scenario = Scenario.GENERAL
    confidence: float = 0.0
    reasoning: str = ""
    # Multi-scenario blend: {scenario: weight} where weights sum to ~1.0
    blend: dict[str, float] = Field(default_factory=dict)
    # Dimension applicability: which dimensions are relevant for this task.
    # Dimensions not in this dict (or mapped to False) will be scored N/A.
    # Empty dict = all dimensions applicable (backward compatibility).
    applicable_dimensions: dict[str, bool] = Field(default_factory=dict)
    dimension_reasoning: dict[str, str] = Field(default_factory=dict)


class EvalSummary(BaseModel):
    """Aggregated evaluation summary."""
    data_content_avg: float | None = None
    structure_usability_avg: float | None = None
    overall_weighted_avg: float | None = None
    dimensions_evaluated: int = 0
    na_dimensions: list[str] = Field(default_factory=list)
    skipped_dimensions: list[str] = Field(default_factory=list)
    weights_applied: dict[str, float] = Field(default_factory=dict)


class CostSummary(BaseModel):
    """Token and cost tracking."""
    total_input_tokens: int = 0
    total_output_tokens: int = 0
    total_latency_ms: int = 0
    total_cost_estimate: float = 0.0
    per_dimension: dict[str, dict[str, float]] = Field(default_factory=dict)


class EvalResult(BaseModel):
    """Complete evaluation result for a single test case."""
    case_id: str
    timestamp: datetime = Field(default_factory=datetime.utcnow)
    excel_file: str
    prompt: str = ""
    prompt_zh: str = ""  # Chinese translation of prompt
    input_files: list[dict[str, Any]] = Field(default_factory=list)
    output_files: list[dict[str, Any]] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)
    scenario: ScenarioResult = Field(default_factory=ScenarioResult)
    dimensions: dict[str, DimensionResult] = Field(default_factory=dict)
    summary: EvalSummary = Field(default_factory=EvalSummary)
    cost: CostSummary = Field(default_factory=CostSummary)
    llm_config: dict[str, Any] = Field(default_factory=dict)


# ── Configuration Models ───────────────────────────────────────────────────


class InputFileConfig(BaseModel):
    """Reference to an input file in the test case."""
    path: str
    description: str = ""


class CaseConfig(BaseModel):
    """Configuration for a single test case (loaded from case.yaml)."""
    id: str
    description: str = ""
    prompt: str
    input_files: list[InputFileConfig] = Field(default_factory=list)
    output_files: list[InputFileConfig] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)
    skip_dimensions: list[str] = Field(default_factory=list)
    scenario: str | None = None  # Manual override


class LLMConfig(BaseModel):
    """LLM provider configuration."""
    provider: str = "anthropic"
    model: str = ""
    api_key: str | None = None
    api_key_env: str | None = None  # Environment variable name
    temperature: float = 0.0
    max_retries: int = 3
    timeout: int = 120
    max_tokens: int = 16384


class EvalConfig(BaseModel):
    """Global evaluation configuration."""
    max_concurrent_calls: int = 8
    screenshot_enabled: bool = True
    max_input_tokens_per_dimension: int = 100000
    language: str = "zh"  # Output language: "zh" (Chinese), "en" (English), etc.


class OutputConfig(BaseModel):
    """Output/reporting configuration."""
    formats: list[str] = Field(default_factory=lambda: ["json", "html", "excel"])
    output_dir: str = "./results"


class GlobalConfig(BaseModel):
    """Top-level configuration combining all settings."""
    llm: LLMConfig = Field(default_factory=LLMConfig)
    evaluation: EvalConfig = Field(default_factory=EvalConfig)
    output: OutputConfig = Field(default_factory=OutputConfig)


# ── Scenario Weights ───────────────────────────────────────────────────────


DEFAULT_WEIGHTS: dict[str, float] = {
    DimensionName.DATA_ACCURACY: 1.0,
    DimensionName.COMPLETENESS: 1.0,
    DimensionName.FORMULA_LOGIC: 1.0,
    DimensionName.RELEVANCE: 1.0,
    DimensionName.SHEET_ORGANIZATION: 1.0,
    DimensionName.TABLE_STRUCTURE: 1.0,
    DimensionName.CHART_APPROPRIATENESS: 1.0,
    DimensionName.PROFESSIONAL_FORMATTING: 0.6,
}

SCENARIO_WEIGHTS: dict[Scenario, dict[str, float]] = {
    # Weights revised based on 62-case GDPVal analysis.
    # Principles:
    #   - DA and Completeness are always important (>=1.0)
    #   - Chart is low everywhere (most workbooks don't have charts)
    #   - PF is always low (formatting matters less than content)
    #   - Relevance is uniform 1.0 (new definition: just "no irrelevant content")
    #   - Scenario-specific boosts only where truly differentiating
    Scenario.REPORTING_ANALYSIS: {
        DimensionName.DATA_ACCURACY: 1.2,
        DimensionName.COMPLETENESS: 1.2,
        DimensionName.FORMULA_LOGIC: 1.0,
        DimensionName.RELEVANCE: 1.0,
        DimensionName.SHEET_ORGANIZATION: 1.0,
        DimensionName.TABLE_STRUCTURE: 1.0,
        DimensionName.CHART_APPROPRIATENESS: 1.0,
        DimensionName.PROFESSIONAL_FORMATTING: 0.6,
    },
    Scenario.DATA_PROCESSING: {
        DimensionName.DATA_ACCURACY: 1.3,
        DimensionName.COMPLETENESS: 1.0,
        DimensionName.FORMULA_LOGIC: 0.8,
        DimensionName.RELEVANCE: 1.0,
        DimensionName.SHEET_ORGANIZATION: 1.0,
        DimensionName.TABLE_STRUCTURE: 1.2,
        DimensionName.CHART_APPROPRIATENESS: 0.5,
        DimensionName.PROFESSIONAL_FORMATTING: 0.5,
    },
    Scenario.TEMPLATE_FORM: {
        DimensionName.DATA_ACCURACY: 1.0,
        DimensionName.COMPLETENESS: 1.0,
        DimensionName.FORMULA_LOGIC: 1.2,
        DimensionName.RELEVANCE: 1.0,
        DimensionName.SHEET_ORGANIZATION: 1.0,
        DimensionName.TABLE_STRUCTURE: 1.3,
        DimensionName.CHART_APPROPRIATENESS: 0.5,
        DimensionName.PROFESSIONAL_FORMATTING: 0.8,
    },
    Scenario.PLANNING_TRACKING: {
        DimensionName.DATA_ACCURACY: 1.0,
        DimensionName.COMPLETENESS: 1.0,
        DimensionName.FORMULA_LOGIC: 1.0,
        DimensionName.RELEVANCE: 1.0,
        DimensionName.SHEET_ORGANIZATION: 1.0,
        DimensionName.TABLE_STRUCTURE: 1.2,
        DimensionName.CHART_APPROPRIATENESS: 0.5,
        DimensionName.PROFESSIONAL_FORMATTING: 0.6,
    },
    Scenario.FINANCIAL_MODELING: {
        DimensionName.DATA_ACCURACY: 1.3,
        DimensionName.COMPLETENESS: 1.2,
        DimensionName.FORMULA_LOGIC: 1.3,
        DimensionName.RELEVANCE: 1.0,
        DimensionName.SHEET_ORGANIZATION: 1.0,
        DimensionName.TABLE_STRUCTURE: 1.0,
        DimensionName.CHART_APPROPRIATENESS: 0.5,
        DimensionName.PROFESSIONAL_FORMATTING: 0.5,
    },
    Scenario.COMPARISON_EVALUATION: {
        DimensionName.DATA_ACCURACY: 1.0,
        DimensionName.COMPLETENESS: 1.2,
        DimensionName.FORMULA_LOGIC: 0.8,
        DimensionName.RELEVANCE: 1.0,
        DimensionName.SHEET_ORGANIZATION: 1.0,
        DimensionName.TABLE_STRUCTURE: 1.2,
        DimensionName.CHART_APPROPRIATENESS: 0.8,
        DimensionName.PROFESSIONAL_FORMATTING: 0.6,
    },
    Scenario.GENERAL: DEFAULT_WEIGHTS,
}


def get_weights(scenario: Scenario) -> dict[str, float]:
    """Get dimension weights for a given scenario."""
    return SCENARIO_WEIGHTS.get(scenario, DEFAULT_WEIGHTS)


def get_blended_weights(scenario_result: ScenarioResult) -> dict[str, float]:
    """Compute blended dimension weights from multi-scenario detection.

    If blend is provided (e.g., {"reporting_analysis": 0.6, "data_processing": 0.4}),
    returns a weighted average of each scenario's dimension weights.
    Falls back to single-scenario weights if no blend is available.
    """
    if not scenario_result.blend:
        return get_weights(scenario_result.detected)

    blended: dict[str, float] = {}
    for dim in DimensionName:
        total = 0.0
        for scenario_str, weight in scenario_result.blend.items():
            try:
                scenario = Scenario(scenario_str)
            except ValueError:
                continue
            dim_weights = SCENARIO_WEIGHTS.get(scenario, DEFAULT_WEIGHTS)
            total += dim_weights.get(dim, 1.0) * weight
        blended[dim] = total

    return blended
