"""Download SpreadsheetBench V2 example cases from HuggingFace.

Usage:
    python scripts/download_ssbv2.py [--output examples]

Downloads the SpreadsheetBench V2 example data from HuggingFace,
creates case.yaml configs, and copies input/output files.
Existing files are skipped (safe to re-run).

Dataset: https://huggingface.co/datasets/KAKA22/SpreadsheetBench-v2
Paper: https://arxiv.org/abs/2406.14991
"""

import argparse
import re
import shutil
import tempfile
import urllib.request
import zipfile
from pathlib import Path


DATASET_URL = (
    "https://huggingface.co/datasets/KAKA22/SpreadsheetBench-v2"
    "/resolve/main/data_example_04_06.zip"
)


def slugify(text: str) -> str:
    """Convert text to a URL/directory-safe slug."""
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    # Collapse multiple hyphens
    slug = re.sub(r"-+", "-", slug)
    return slug


def read_prompt(prompt_path: Path) -> str:
    """Read a prompt.txt file, return stripped content."""
    if not prompt_path.exists():
        return ""
    return prompt_path.read_text(encoding="utf-8").strip()


def copy_if_missing(src: Path, dst: Path) -> bool:
    """Copy file if destination doesn't exist. Returns True if copied."""
    if dst.exists():
        return False
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return True


def write_case_yaml(case_dir: Path, case_data: dict) -> None:
    """Write case.yaml if it doesn't exist."""
    yaml_path = case_dir / "case.yaml"
    if yaml_path.exists():
        return

    import yaml

    with open(yaml_path, "w", encoding="utf-8") as f:
        yaml.dump(case_data, f, allow_unicode=True,
                  default_flow_style=False, sort_keys=False)


def process_debugging(data_root: Path, output_base: Path) -> tuple[int, int]:
    """Process Debugging category: shared prompt, per-task input/golden."""
    debug_dir = data_root / "Debugging"
    if not debug_dir.exists():
        return 0, 0

    shared_prompt = read_prompt(debug_dir / "prompt.txt")
    cases = 0
    files_copied = 0

    for task_dir in sorted(debug_dir.iterdir()):
        if not task_dir.is_dir():
            continue

        # Find input and golden files
        input_files = list(task_dir.glob("*_input.xlsx"))
        golden_files = list(task_dir.glob("*_golden.xlsx"))
        if not input_files or not golden_files:
            continue

        input_file = input_files[0]
        golden_file = golden_files[0]

        # Extract error type from filename: "04_06_Inconsistent Color Coding_input.xlsx"
        name_part = input_file.stem.replace("_input", "")
        # Remove the leading ID prefix (e.g., "04_06_")
        error_type = re.sub(r"^\d+_\d+_", "", name_part)
        task_id = task_dir.name

        case_name = f"ssb2-debug-{slugify(task_id)}-{slugify(error_type)}"
        case_dir = output_base / case_name
        case_dir.mkdir(parents=True, exist_ok=True)
        (case_dir / "input").mkdir(exist_ok=True)
        (case_dir / "output").mkdir(exist_ok=True)

        if copy_if_missing(input_file, case_dir / "input" / input_file.name):
            files_copied += 1
        if copy_if_missing(golden_file, case_dir / "output" / golden_file.name):
            files_copied += 1

        write_case_yaml(case_dir, {
            "id": case_name,
            "description": f"SpreadsheetBench V2 Debugging - {error_type}",
            "prompt": shared_prompt,
            "input_files": [
                {"path": f"input/{input_file.name}",
                 "description": "Spreadsheet with errors to fix"},
            ],
            "output_files": [
                {"path": f"output/{golden_file.name}",
                 "description": "Corrected spreadsheet"},
            ],
            "metadata": {
                "source": "KAKA22/SpreadsheetBench-v2",
                "category": "debugging",
                "task_id": task_id,
                "error_type": error_type,
            },
        })
        cases += 1

    return cases, files_copied


def process_template(data_root: Path, output_base: Path) -> tuple[int, int]:
    """Process Template category: per-task prompt + input/golden."""
    tmpl_dir = data_root / "Template"
    if not tmpl_dir.exists():
        return 0, 0

    cases = 0
    files_copied = 0

    for task_dir in sorted(tmpl_dir.iterdir()):
        if not task_dir.is_dir():
            continue

        prompt = read_prompt(task_dir / "prompt.txt")
        input_files = list(task_dir.glob("*_input.xlsx"))
        golden_files = list(task_dir.glob("*_golden.xlsx"))
        if not input_files or not golden_files:
            continue

        input_file = input_files[0]
        golden_file = golden_files[0]
        task_id = task_dir.name

        case_name = f"ssb2-template-{slugify(task_id)}"
        case_dir = output_base / case_name
        case_dir.mkdir(parents=True, exist_ok=True)
        (case_dir / "input").mkdir(exist_ok=True)
        (case_dir / "output").mkdir(exist_ok=True)

        if copy_if_missing(input_file, case_dir / "input" / input_file.name):
            files_copied += 1
        if copy_if_missing(golden_file, case_dir / "output" / golden_file.name):
            files_copied += 1

        write_case_yaml(case_dir, {
            "id": case_name,
            "description": f"SpreadsheetBench V2 Template - {task_id}",
            "prompt": prompt,
            "input_files": [
                {"path": f"input/{input_file.name}",
                 "description": "Template spreadsheet to complete"},
            ],
            "output_files": [
                {"path": f"output/{golden_file.name}",
                 "description": "Completed template"},
            ],
            "metadata": {
                "source": "KAKA22/SpreadsheetBench-v2",
                "category": "template",
                "task_id": task_id,
            },
        })
        cases += 1

    return cases, files_copied


def process_visualization(data_root: Path, output_base: Path) -> tuple[int, int]:
    """Process Visualization category: per-task prompt + input/golden."""
    viz_dir = data_root / "Visualization"
    if not viz_dir.exists():
        return 0, 0

    cases = 0
    files_copied = 0

    for task_dir in sorted(viz_dir.iterdir()):
        if not task_dir.is_dir():
            continue

        prompt = read_prompt(task_dir / "prompt.txt")
        input_files = list(task_dir.glob("*_input.xlsx"))
        golden_files = list(task_dir.glob("*_golden.xlsx"))
        if not input_files or not golden_files:
            continue

        input_file = input_files[0]
        golden_file = golden_files[0]
        task_name = task_dir.name  # e.g., "Task 126"

        case_name = f"ssb2-viz-{slugify(task_name)}"
        case_dir = output_base / case_name
        case_dir.mkdir(parents=True, exist_ok=True)
        (case_dir / "input").mkdir(exist_ok=True)
        (case_dir / "output").mkdir(exist_ok=True)

        if copy_if_missing(input_file, case_dir / "input" / input_file.name):
            files_copied += 1
        if copy_if_missing(golden_file, case_dir / "output" / golden_file.name):
            files_copied += 1

        write_case_yaml(case_dir, {
            "id": case_name,
            "description": f"SpreadsheetBench V2 Visualization - {task_name}",
            "prompt": prompt,
            "input_files": [
                {"path": f"input/{input_file.name}",
                 "description": "Spreadsheet with data for visualization"},
            ],
            "output_files": [
                {"path": f"output/{golden_file.name}",
                 "description": "Spreadsheet with expected charts/visualizations"},
            ],
            "metadata": {
                "source": "KAKA22/SpreadsheetBench-v2",
                "category": "visualization",
                "task_id": task_name,
            },
        })
        cases += 1

    return cases, files_copied


def process_financial_model(data_root: Path, output_base: Path) -> tuple[int, int]:
    """Process Financial_Model category.

    Each project folder has multiple steps (increasing difficulty) sharing
    one golden answer. We create one case per step. Only the final step
    (highest number) asks to complete ALL tabs, matching the golden fully.
    Earlier steps are partial — the golden still works as reference but the
    evaluator should note that only specific tabs were requested.
    """
    fm_dir = data_root / "Financial_Model"
    if not fm_dir.exists():
        return 0, 0

    cases = 0
    files_copied = 0

    for project_dir in sorted(fm_dir.iterdir()):
        if not project_dir.is_dir():
            continue

        # Find golden file (pattern: XX_Company_golden.xlsx)
        golden_files = list(project_dir.glob("*_golden.xlsx"))
        if not golden_files:
            continue
        golden_file = golden_files[0]

        # Parse company name from golden: "09_PepsiCo_golden.xlsx" → "PepsiCo"
        golden_stem = golden_file.stem.replace("_golden", "")
        parts = golden_stem.split("_", 1)
        company_name = parts[1] if len(parts) > 1 else parts[0]

        # Find step files: XX_NN_Company_input.xlsx + XX_NN_Company_prompt.txt
        step_inputs = sorted(project_dir.glob("*_input.xlsx"))
        for input_file in step_inputs:
            # Parse step number: "09_03_PepsiCo_input.xlsx" → step 3
            stem = input_file.stem.replace("_input", "")
            match = re.match(r"\d+_(\d+)_", stem)
            if not match:
                continue
            step_num = int(match.group(1))

            # Find matching prompt
            prompt_file = input_file.with_name(
                input_file.name.replace("_input.xlsx", "_prompt.txt")
            )
            prompt = read_prompt(prompt_file)
            if not prompt:
                continue

            case_name = f"ssb2-fm-{slugify(company_name)}-step{step_num}"
            case_dir = output_base / case_name
            case_dir.mkdir(parents=True, exist_ok=True)
            (case_dir / "input").mkdir(exist_ok=True)
            (case_dir / "output").mkdir(exist_ok=True)

            if copy_if_missing(input_file, case_dir / "input" / input_file.name):
                files_copied += 1
            if copy_if_missing(golden_file, case_dir / "output" / golden_file.name):
                files_copied += 1

            write_case_yaml(case_dir, {
                "id": case_name,
                "description": (
                    f"SpreadsheetBench V2 Financial Model"
                    f" - {company_name} Step {step_num}"
                ),
                "prompt": prompt,
                "input_files": [
                    {"path": f"input/{input_file.name}",
                     "description": f"Financial model workbook (step {step_num})"},
                ],
                "output_files": [
                    {"path": f"output/{golden_file.name}",
                     "description": "Completed financial model"},
                ],
                "metadata": {
                    "source": "KAKA22/SpreadsheetBench-v2",
                    "category": "financial_model",
                    "company": company_name,
                    "step": step_num,
                    "total_steps": len(step_inputs),
                },
            })
            cases += 1

    return cases, files_copied


def main():
    parser = argparse.ArgumentParser(
        description="Download SpreadsheetBench V2 example cases"
    )
    parser.add_argument(
        "--output", default="examples",
        help="Output directory for cases (default: examples)",
    )
    args = parser.parse_args()

    output_base = Path(args.output)

    # Download and extract
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        zip_path = tmp_path / "ssbv2.zip"

        print("Downloading SpreadsheetBench V2 from HuggingFace...")
        urllib.request.urlretrieve(DATASET_URL, str(zip_path))
        print("Extracting...")

        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(tmp_path)

        # Find the extracted root (may be nested in a folder)
        extracted_dirs = [
            d for d in tmp_path.iterdir()
            if d.is_dir() and not d.name.startswith(".")
        ]
        if not extracted_dirs:
            print("ERROR: No data found in downloaded archive")
            return
        data_root = extracted_dirs[0]

        print(f"Processing cases from: {data_root.name}")

        total_cases = 0
        total_files = 0

        # Process each category
        for name, processor in [
            ("Debugging", process_debugging),
            ("Template", process_template),
            ("Visualization", process_visualization),
            ("Financial Model", process_financial_model),
        ]:
            n_cases, n_files = processor(data_root, output_base)
            if n_cases:
                print(f"  {name}: {n_cases} cases, {n_files} files copied")
            total_cases += n_cases
            total_files += n_files

    print(f"\nDone: {total_cases} cases, {total_files} files copied")
    print(f"Cases are in: {output_base.resolve()}")


if __name__ == "__main__":
    main()
