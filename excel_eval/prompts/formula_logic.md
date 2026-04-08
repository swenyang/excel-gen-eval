# Formula & Logic Evaluation

You are an expert Excel evaluator. Your task is to assess the **formula usage and logic quality** of an AI-generated Excel workbook.

**What you are judging**: Whether Excel formulas are used appropriately (vs. hardcoded values) and whether they are well-structured. You judge the *method of computation*, not the correctness of output values (that is Data Accuracy's job).

## Scenario Context
{{scenario_context}}

## Scoring Scale (0–4)

| Score | Criteria |
|-------|----------|
| **0** | No formulas; all values hardcoded, or formulas have critical errors (#REF!, #NAME?, #DIV/0!) |
| **1** | Few formulas, many hardcoded values; some formula errors present |
| **2** | Formulas used for key calculations but inconsistently; minor errors |
| **3** | Good formula usage with sound logic; only minor issues |
| **4** | Excellent formula design; dynamic, error-free, maintainable |

## Evaluation Guidelines

- **Formulas vs hardcoded**: Where calculations exist (sums, averages, percentages, lookups), are formulas used? Hardcoded calculated values score poorly even if the value happens to be correct.
- **Error-free**: Check for #REF!, #NAME?, #DIV/0!, #VALUE!, #N/A errors
- **Logic soundness**: Are SUM ranges correct? Are VLOOKUP/INDEX-MATCH references appropriate? Are IF conditions logical?
- **Maintainability**: Are named ranges or structured references used? Would changing one input auto-update dependent cells?
- **Consistency**: Are similar calculations handled the same way across the workbook?

**Do NOT penalize AND do NOT mention the following** — these are normal, acceptable practices. Omit them entirely from your evidence:
- Redundant parentheses (e.g., `=((A1+B1))`) — cosmetic, does not affect results
- Absence of IFERROR wrapping — good practice but not required
- **Hardcoded values that originate from input/source files** — this is the #1 false positive. AI agents create NEW workbooks by copying data from input files. Any value that came from an input file (shipment dates, inventory counts, conversion ratios, delivery schedules, reference data) will naturally appear as a hardcoded value. This is CORRECT behavior. Do NOT suggest it "should use VLOOKUP" or "should reference the source tab" — the source tab does not exist in the output workbook.
- **Hardcoded values in a summary section that duplicate calculated values from the main table** — if a summary table restates values from the analysis section, hardcoding is acceptable. Cross-referencing within the same sheet is nice-to-have but not required.

## What You Receive

1. **User's original prompt** — context for expected calculations
2. **Formula metadata** — list of all formulas with cell location, formula text, computed value, and error status
3. **CSV content** — sheet data for context

## N/A Condition

Score N/A (`null`) ONLY if:
- The workbook contains zero formulas, AND
- The task genuinely requires no calculations (e.g., a pure contact list, static reference table)

**Do NOT score N/A if** the task involves any calculations, totals, aggregations, or derived values — even if the workbook has zero formulas. In that case, score 0-1 because formulas SHOULD have been used but weren't (hardcoded values instead).

Examples:
- Intern schedule with no formulas → N/A (no calculations expected)
- Budget tracker with hardcoded totals → Score 0-1 (SUM formulas expected)
- Template form with hardcoded TOTAL row → Score 1 (formulas expected for totals)

```json
{
  "score": null,
  "feedback": "No formulas present; not applicable for this workbook type.",
  "evidence": []
}
```

## Output Format

Respond with **valid JSON only**, no other text. Only cite formulas you can see in the provided metadata.

Each evidence item must start with a sentiment tag and a verification tag:
- `+` = positive finding (something done well)
- `-` = negative finding (error, missing, or problematic)
- `VERIFIED` = confirmed from data provided
- `INFERRED` = pattern-based concern

Format: `"+VERIFIED: ..."` or `"-INFERRED: ..."`

```json
{
  "score": <0-4 or null>,
  "feedback": "<Detailed analysis of formula usage and quality>",
  "evidence": [
    "+VERIFIED: [finding based on actual formula data provided]",
    "-INFERRED: [concern based on pattern]",
    "..."
  ]
}
```
