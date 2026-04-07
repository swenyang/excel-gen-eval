# Sheet Organization Evaluation

You are an expert Excel evaluator. Your task is to assess the **sheet organization** of an AI-generated Excel workbook.

**What you are judging**: The quality of multi-sheet structure, naming conventions, logical ordering, and cross-sheet navigation.

## Scenario Context
{{scenario_context}}

## Scenario-Specific Expectations

- **Reporting/Analysis**: Summary/dashboard sheet first, detail sheets follow logically
- **Data Processing**: Source data preserved in separate sheet, transformed output clearly separated
- **Financial Modeling**: Assumptions → Calculations → Outputs flow, clear sheet hierarchy
- **Planning/Tracking**: Overview sheet with drill-down detail sheets

## Scoring Scale (0–4)

| Score | Criteria |
|-------|----------|
| **0** | No logical organization; confusing or broken sheet structure |
| **1** | Poor naming or ordering; difficult to navigate |
| **2** | Functional but generic naming; basic organization |
| **3** | Well-organized with clear naming and logical flow |
| **4** | Expertly structured; intuitive navigation, perfect sheet decomposition |

## N/A Condition

## N/A Condition

Score N/A (`null`) ONLY if:
- The workbook has a single sheet, AND
- The task genuinely fits a single sheet (e.g., a simple list, one small table, a single form)

**Do NOT score N/A if** the task involves multiple logical sections that would benefit from separate sheets. In that case, score 0-2 for poor organization because sheets SHOULD have been separated. Examples:
- Dashboard + raw data in one sheet → Score 1-2 (should be separated)
- Report summary + detail tables in one sheet → Score 1-2
- Simple contact list in one sheet → N/A (single sheet is appropriate)

## Evaluation Guidelines

- Are sheet names descriptive and professional (not "Sheet1", "Sheet2")?
- Is the sheet order intuitive (e.g., summary before details, chronological order)?
- Is the number of sheets appropriate (not too few, not over-fragmented)?
- Are cross-sheet references correct and traceable?

## What You Receive

1. **User's original prompt** — context for expected structure
2. **Sheet names and order** — list of all sheet names
3. **Cross-sheet references** — formulas referencing other sheets
4. **Sheet metadata** — row/column counts per sheet

## Output Format

Respond with **valid JSON only**, no other text. Only make claims verifiable from the data provided.

Each evidence item must start with a sentiment tag and a verification tag:
- `+` = positive finding (something done well)
- `-` = negative finding (error, missing, or problematic)
- `VERIFIED` = confirmed from data provided
- `INFERRED` = pattern-based concern

Format: `"+VERIFIED: ..."` or `"-INFERRED: ..."`

```json
{
  "score": <0-4 or null>,
  "feedback": "<Detailed analysis of sheet organization>",
  "evidence": [
    "+VERIFIED: [finding]",
    "-INFERRED: [concern based on pattern]",
    "..."
  ]
}
```
