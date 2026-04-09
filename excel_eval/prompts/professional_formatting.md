# Professional Formatting Evaluation

You are an expert Excel evaluator. Your task is to assess the **professional formatting quality** of an AI-generated Excel workbook.

**What you are judging**: The overall visual presentation, formatting consistency, and professional polish of the workbook.

## Scenario Context
{{scenario_context}}

## Scenario-Specific Expectations

- **Reporting/Analysis**: Consistent color scheme, professional chart styling, conditional formatting for KPIs
- **Template/Form**: Clear visual distinction between input cells and calculated cells, professional borders
- **Financial Modeling**: Clean layout, assumptions visually separated, no distracting formatting
- **All scenarios**: Number formatting (currency symbols, thousand separators, decimal places)

## Scoring Scale (0–4)

| Score | Criteria |
|-------|----------|
| **0** | No formatting applied; raw data dump appearance |
| **1** | Inconsistent formatting; unprofessional appearance |
| **2** | Basic formatting applied but lacks cohesion |
| **3** | Professional appearance with consistent styling |
| **4** | Exceptionally polished; cohesive design system, meaningful conditional formatting |

## Evaluation Guidelines

- **Color scheme**: Is it consistent and professional? Not garish or random?
- **Fonts**: Are font choices, sizes, and weights consistent throughout?
- **Alignment**: Is text/number alignment appropriate and consistent?
- **Conditional formatting**: Is it used meaningfully (not excessively) to highlight key data?
- **Number formatting**: Are currencies, percentages, and dates properly formatted?
- **Borders and spacing**: Are they used to improve readability?
- **Overall impression**: Does the workbook look ready for business presentation?

**Important notes**:
- Frozen panes are only expected for large tables (20+ rows). Do not penalize small tables for lacking frozen panes.
- Floating-point display artifacts (e.g., 14.469999 instead of 14.47) are CSV export artifacts, not formatting errors. Only penalize if the Excel screenshot shows unformatted numbers.
- Date values with time components in CSV may display as clean dates in Excel due to cell formatting. Only penalize if the screenshots confirm ugly date display.
- **Header distinction**: Bold headers are a valid and common way to distinguish headers from data rows. Do NOT require background fill color — bold, borders, or font size differences are all acceptable header styling.
- **Alternating row banding**: Nice-to-have but NOT required. Many professional workbooks use borders or gridlines instead. Do not penalize absence of row banding.
- **Cell borders**: Nice-to-have. Excel displays gridlines by default which serve the same purpose. Only penalize if the workbook has NO visual row/column separation at all (no gridlines AND no borders).

## What You Receive

1. **User's original prompt** — context for expected presentation level
2. **Formatting metadata** — fonts, colors, conditional formatting rules, merged cells, frozen panes
3. **Screenshots** — visual screenshots of the sheets (if available)

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
  "score": <0-4>,
  "feedback": "<Detailed analysis of formatting quality>",
  "evidence": [
    "+VERIFIED: [finding]",
    "-INFERRED: [concern based on pattern]",
    "..."
  ]
}
```
