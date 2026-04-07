"""Dimension evaluators and scenario detection."""

from excel_eval.evaluators.base import BaseEvaluator
from excel_eval.evaluators.scenario_detector import ScenarioDetector
from excel_eval.evaluators.data_accuracy import DataAccuracyEvaluator
from excel_eval.evaluators.completeness import CompletenessEvaluator
from excel_eval.evaluators.formula_logic import FormulaLogicEvaluator
from excel_eval.evaluators.relevance import RelevanceEvaluator
from excel_eval.evaluators.sheet_organization import SheetOrganizationEvaluator
from excel_eval.evaluators.table_structure import TableStructureEvaluator
from excel_eval.evaluators.chart_appropriateness import ChartAppropriatenessEvaluator
from excel_eval.evaluators.professional_formatting import ProfessionalFormattingEvaluator

ALL_EVALUATORS: list[type[BaseEvaluator]] = [
    DataAccuracyEvaluator,
    CompletenessEvaluator,
    FormulaLogicEvaluator,
    RelevanceEvaluator,
    SheetOrganizationEvaluator,
    TableStructureEvaluator,
    ChartAppropriatenessEvaluator,
    ProfessionalFormattingEvaluator,
]

__all__ = [
    "BaseEvaluator",
    "ScenarioDetector",
    "DataAccuracyEvaluator",
    "CompletenessEvaluator",
    "FormulaLogicEvaluator",
    "RelevanceEvaluator",
    "SheetOrganizationEvaluator",
    "TableStructureEvaluator",
    "ChartAppropriatenessEvaluator",
    "ProfessionalFormattingEvaluator",
    "ALL_EVALUATORS",
]
