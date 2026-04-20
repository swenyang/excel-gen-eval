"""JSON schemas for Anthropic structured output (output_config)."""

# Dimension evaluator response: {score, feedback, evidence[]}
DIMENSION_EVAL_SCHEMA: dict = {
    "type": "object",
    "properties": {
        "score": {
            "type": "integer",
            "description": "Score from 0 to 4",
        },
        "feedback": {
            "type": "string",
            "description": "Detailed analysis with numbered findings",
        },
        "evidence": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Evidence items prefixed with +/-VERIFIED/INFERRED tags",
        },
    },
    "required": ["score", "feedback", "evidence"],
    "additionalProperties": False,
}

# Scenario detection response: {scenario, confidence, reasoning, blend, applicable_dimensions}
SCENARIO_DETECT_SCHEMA: dict = {
    "type": "object",
    "properties": {
        "scenario": {
            "type": "string",
            "description": "Primary scenario key",
        },
        "confidence": {
            "type": "number",
            "description": "Confidence score from 0.0 to 1.0",
        },
        "reasoning": {
            "type": "string",
            "description": "Brief explanation (1-3 sentences)",
        },
        "blend": {
            "type": "object",
            "properties": {
                "reporting_analysis": {"type": "number"},
                "data_processing": {"type": "number"},
                "template_form": {"type": "number"},
                "planning_tracking": {"type": "number"},
                "financial_modeling": {"type": "number"},
                "comparison_evaluation": {"type": "number"},
                "general": {"type": "number"},
            },
            "additionalProperties": False,
            "description": "Scenario blend weights summing to 1.0",
        },
        "applicable_dimensions": {
            "type": "object",
            "properties": {
                "data_accuracy": {"type": "boolean"},
                "completeness": {"type": "boolean"},
                "formula_logic": {"type": "boolean"},
                "relevance": {"type": "boolean"},
                "sheet_organization": {"type": "boolean"},
                "table_structure": {"type": "boolean"},
                "chart_appropriateness": {"type": "boolean"},
                "professional_formatting": {"type": "boolean"},
            },
            "additionalProperties": False,
            "description": "Which dimensions are applicable for this task",
        },
        "dimension_reasoning": {
            "type": "object",
            "properties": {
                "data_accuracy": {"type": "string"},
                "completeness": {"type": "string"},
                "formula_logic": {"type": "string"},
                "relevance": {"type": "string"},
                "sheet_organization": {"type": "string"},
                "table_structure": {"type": "string"},
                "chart_appropriateness": {"type": "string"},
                "professional_formatting": {"type": "string"},
            },
            "additionalProperties": False,
            "description": "Brief reasoning for dimensions set to false",
        },
    },
    "required": ["scenario", "confidence", "reasoning", "blend",
                  "applicable_dimensions"],
    "additionalProperties": False,
}

# Evidence verification response: [{index, confirmed, reason}]
EVIDENCE_VERIFY_SCHEMA: dict = {
    "type": "object",
    "properties": {
        "results": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "index": {
                        "type": "integer",
                        "description": "1-based claim index",
                    },
                    "confirmed": {
                        "type": "boolean",
                        "description": "Whether the claim is confirmed",
                    },
                    "reason": {
                        "type": "string",
                        "description": "Brief explanation",
                    },
                },
                "required": ["index", "confirmed", "reason"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["results"],
    "additionalProperties": False,
}
