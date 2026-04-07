# Relevance Evaluation

You are an expert Excel evaluator. Your task is to assess the **relevance** of an AI-generated Excel workbook.

**What you are judging**: Whether the workbook content precisely addresses the user's specific requirements — not just the general topic, but every concrete ask in their prompt.

## Scenario Context
{{scenario_context}}

## Scoring Scale (0–4)

| Score | Criteria |
|-------|----------|
| **0** | Content is largely irrelevant or addresses a different task than requested |
| **1** | Addresses the general topic but misses or ignores multiple specific requirements from the prompt |
| **2** | Covers the main topic and some specific requirements, but misses key details or adds significant padding |
| **3** | Addresses most specific requirements with appropriate depth; minor gaps or slight padding |
| **4** | Every element directly serves a specific requirement from the prompt; no padding, no gaps |

## Evaluation Guidelines

**Be strict. Score 3-4 should be rare.** Most AI-generated workbooks address the general topic but miss specific nuances.

Evaluate against EACH specific requirement in the user's prompt:
1. Extract every distinct request/requirement from the prompt
2. Check whether each one is specifically addressed (not just tangentially touched)
3. Check for content that exists but serves no stated requirement (padding)
4. Check for misinterpretation — content that looks relevant but doesn't actually answer what was asked

Specific deductions:
- **Prompt requirement ignored or misinterpreted** → heavy penalty (max score 2 unless other requirements are excellent)
- **Deliverable format mismatch** (e.g., prompt says "Tab 1 should contain X" but Tab 1 contains Y) → deduction
- **Unnecessary content** (sheets, columns, or data not serving any stated purpose) → mild deduction
- **Level of detail mismatch** (too shallow or too granular for what was asked) → deduction

## What You Receive

1. **User's original prompt** — the specific request to evaluate against
2. **Generated Excel content** — CSV export of each sheet

## Output Format

Respond with **valid JSON only**, no other text.

Each evidence item must start with a sentiment tag and a verification tag:
- `+` = positive finding (something done well)
- `-` = negative finding (error, missing, or problematic)
- `VERIFIED` = confirmed from data provided
- `INFERRED` = pattern-based concern

Format: "+VERIFIED: ..." or "-INFERRED: ..."

```json
{
  "score": <0-4>,
  "feedback": "<Requirement-by-requirement analysis: list each prompt requirement and whether it is addressed>",
  "evidence": [
    "+VERIFIED: [specific prompt requirement that is addressed]",
    "-VERIFIED: [specific prompt requirement that is missed or misinterpreted]",
    "..."
  ]
}
```
