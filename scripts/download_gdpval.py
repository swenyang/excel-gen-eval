"""Download all GDPVal Excel cases from HuggingFace.

Usage:
    python scripts/download_gdpval.py [--output examples]

Downloads the 62 Excel-deliverable cases from openai/gdpval dataset,
creates case.yaml configs, and downloads input/output files.
Existing files are skipped (safe to re-run).
"""

import argparse
import re
import urllib.request
from pathlib import Path


# Semantic names for cases (hand-curated for readability)
CASE_NAMES = {
    "83d10b06": "gdpval-audit-sample",
    "7b08cd4d": "accounting-music-tour",
    "7d7fc9a7": "accounting-amortization",
    "ee09d943": "accounting-financials",
    "17111c03": "planning-cleanup-schedule",
    "c44e9b62": "government-fte-report",
    "24d1e93f": "financial-npv-headlamp",
    "7bbfcfe9": "compliance-scra-questions",
    "dfb4e0cd": "compliance-grants-analysis",
    "4c18ebae": "data-processing-transactions",
    "c357f0e2": "it-uat-plan",
    "b7a5912e": "rental-daily-report",
    "aa071045": "rental-damage-report",
    "b39a5aa7": "finance-orchestra-comp",
    "4520f882": "finance-theatre-cba",
    "62f04c2f": "template-exchange-form",
    "3f821c2d": "wholesale-stock-final",
    "e996036e": "wholesale-scenario-planning",
    "327fbc21": "wholesale-sales-plan",
    "6dcae3f5": "healthcare-key-indicators",
    "40a8c4b1": "healthcare-grand-rounds",
    "bf68f2ad": "manufacturing-catchup-plan",
    "efca245f": "manufacturing-recovery-plan",
    "9e39df84": "manufacturing-dashboard",
    "68d8d901": "manufacturing-freeze-dry-plan",
    "1752cb53": "manufacturing-wire-test-plan",
    "0fad6023": "retail-meat-pog-template",
    "4d61a19a": "retail-promotion-template",
    "40a99a31": "manufacturing-cnc-interface",
    "5a2d70da": "manufacturing-cover-plate",
    "81db15ff": "healthcare-np-pa-allowances",
    "61e7b9c6": "healthcare-menopause-formulary",
    "41f6ef59": "healthcare-declined-payments",
    "a0552909": "healthcare-bulk-forms",
    "4b98ccce": "healthcare-patient-incident",
    "b5d2e6f1": "wholesale-sales-analysis",
    "f841ddcf": "wholesale-po-log",
    "47ef842d": "wholesale-inventory",
    "1137e2bb": "wholesale-po-audit",
    "c3525d4d": "wholesale-holiday-budget",
    "c657103b": "finance-roth-conversion",
    "a079d38f": "media-cost-breakdown",
    "02aa1805": "project-water-wells",
    "ce864f41": "project-work-decomposition",
    "a99d85fc": "realestate-rent-matrix",
    "650adcb1": "general-intern-schedule",
    "a73fbc98": "recreation-spring-bazaar",
    "dd724c67": "nursing-transition-care",
    "90edba97": "nursing-lab-tracker",
    "f2986c1f": "pharmacy-drug-list",
    "d7cfae6f": "wholesale-risks-selling",
    "19403010": "reporting-sales-analysis",
    "7ed932dd": "wholesale-additional-shipments",
    "105f8ad0": "comparison-pricing-strategy",
    "b57efde3": "wholesale-prospect-list",
    "15d37511": "wholesale-uv-pro-forma",
    "bb863dd9": "wholesale-medical-quotation",
    "6a900a40": "wholesale-transport-quotation",
    "1d4672c8": "finance-correlation-matrix",
    "5349dd7b": "shipping-rate-analysis",
    "11dcc268": "shipping-location-report",
    "76418a2c": "shipping-daily-manifest",
}


def get_case_name(task_id: str, occupation: str) -> str:
    """Get semantic case name, falling back to auto-generated."""
    short = task_id[:8]
    if short in CASE_NAMES:
        return CASE_NAMES[short]
    occ = re.sub(r"[^a-z0-9]+", "-", occupation.lower())[:30].strip("-")
    return f"gdpval-{short}-{occ}"


def download_file(url: str, dest: Path) -> bool:
    """Download a file if it doesn't exist. Returns True if downloaded."""
    if dest.exists():
        return False
    try:
        urllib.request.urlretrieve(url, str(dest))
        return True
    except Exception as e:
        print(f"  WARNING: Failed to download {dest.name}: {e}")
        return False


def write_case_yaml(case_dir: Path, case_name: str, row: dict,
                     ref_files: list, del_xlsx: list) -> None:
    """Write case.yaml if it doesn't exist."""
    yaml_path = case_dir / "case.yaml"
    if yaml_path.exists():
        return

    lines = [
        f"id: {case_name}",
        f'description: "GDPVal {row[\"task_id\"][:8]} - {row[\"occupation\"]}"',
        "",
        "prompt: |",
    ]
    for line in row["prompt"].split("\n"):
        lines.append(f"  {line}")

    lines += ["", "input_files:"]
    for f in ref_files:
        fname = f.split("/")[-1]
        lines.append(f'  - path: "input/{fname}"')
        lines.append(f'    description: "{fname}"')
    if not ref_files:
        lines.append("  []")

    lines += ["", "output_files:"]
    for f in del_xlsx:
        fname = f.split("/")[-1]
        lines.append(f'  - path: "output/{fname}"')
        lines.append(f'    description: "{fname}"')

    lines += [
        "",
        "metadata:",
        f"  source: openai/gdpval",
        f'  task_id: "{row["task_id"]}"',
        f'  sector: "{row["sector"]}"',
        f'  occupation: "{row["occupation"]}"',
    ]

    yaml_path.write_text("\n".join(lines), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Download GDPVal Excel cases")
    parser.add_argument("--output", default="examples", help="Output directory")
    args = parser.parse_args()

    try:
        from datasets import load_dataset
    except ImportError:
        print("ERROR: 'datasets' package required. Install with: pip install datasets")
        return

    print("Loading GDPVal dataset from HuggingFace...")
    ds = load_dataset("openai/gdpval", split="train")

    base = Path(args.output)
    total = 0
    downloaded = 0

    for row in ds:
        deliverables = row.get("deliverable_files", [])
        has_excel = any(f.lower().endswith((".xlsx", ".xls", ".xlsm"))
                        for f in deliverables)
        if not has_excel:
            continue

        total += 1
        case_name = get_case_name(row["task_id"], row["occupation"])
        case_dir = base / case_name
        (case_dir / "input").mkdir(parents=True, exist_ok=True)
        (case_dir / "output").mkdir(parents=True, exist_ok=True)

        # Download input files
        ref_urls = row.get("reference_file_urls", [])
        ref_files = row.get("reference_files", [])
        for url, fpath in zip(ref_urls, ref_files):
            fname = fpath.split("/")[-1]
            if download_file(url, case_dir / "input" / fname):
                downloaded += 1

        # Download output Excel files
        del_urls = row.get("deliverable_file_urls", [])
        del_xlsx = [f for f in deliverables
                     if f.lower().endswith((".xlsx", ".xls", ".xlsm"))]
        del_xlsx_urls = [u for u, f in zip(del_urls, deliverables)
                          if f.lower().endswith((".xlsx", ".xls", ".xlsm"))]
        for url, fpath in zip(del_xlsx_urls, del_xlsx):
            fname = fpath.split("/")[-1]
            if download_file(url, case_dir / "output" / fname):
                downloaded += 1

        # Write case.yaml
        write_case_yaml(case_dir, case_name, row, ref_files, del_xlsx)

    print(f"\nDone: {total} cases, {downloaded} files downloaded")
    print(f"Cases are in: {base.resolve()}")


if __name__ == "__main__":
    main()
