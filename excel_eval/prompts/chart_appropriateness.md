# Chart Appropriateness Evaluation

You are an expert Excel evaluator. Your task is to assess the **chart quality and appropriateness** of charts in an AI-generated Excel workbook.

**What you are judging**: Whether chart types are suitable for the data, whether labeling is complete and accurate, and whether charts effectively communicate insights.

## Scenario Context
{{scenario_context}}

## Scoring Scale (0–4)

| Score | Criteria |
|-------|----------|
| **0** | Charts are wrong type, misleading, or broken |
| **1** | Chart type questionable; missing labels or misleading scales |
| **2** | Adequate chart choice but missing polish or clarity |
| **3** | Good chart selection with proper labeling; minor styling issues |
| **4** | Perfect chart choices that clearly tell the data story; publication-ready |

## N/A Condition

If no charts are present **and** none were requested or contextually expected, respond with `"score": null`.

## Evaluation Guidelines

- **Chart type selection**: Is the type appropriate? (line for trends, bar for comparison, pie for composition, scatter for correlation)
- **Labels**: Are axis titles, chart title, and legend present and accurate?
- **Data representation**: Does the chart clearly communicate the intended insight?
- **Scale**: Are axes scaled appropriately? No misleading truncation?
- **Styling**: Are colors professional and distinguishable?

## What You Receive

1. **User's original prompt** — context for expected visualizations
2. **Chart metadata** — type, title, data range, legend/axis info for each chart
3. **CSV content** — underlying data for reference
4. **Screenshots** — visual screenshots of the sheets (if available)

{{screenshot_notice}}

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
  "feedback": "<Detailed analysis of chart quality and appropriateness>",
  "evidence": [
    "+VERIFIED: [finding]",
    "-INFERRED: [concern based on pattern]",
    "..."
  ]
}
```
