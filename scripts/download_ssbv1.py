"""Download SpreadsheetBench V1 example cases from HuggingFace.

Usage:
    python scripts/download_ssbv1.py [--output examples_ssbv1]

Downloads the SpreadsheetBench V1 verified 400 cases from HuggingFace,
creates case.yaml configs, and copies input/output files.
Existing files are skipped (safe to re-run).

Dataset: https://huggingface.co/datasets/KAKA22/SpreadsheetBench
Paper: https://arxiv.org/abs/2406.14991
"""

import argparse
import re
import shutil
import tarfile
import tempfile
import urllib.request
from pathlib import Path


DATASET_URL = (
    "https://huggingface.co/datasets/KAKA22/SpreadsheetBench"
    "/resolve/main/spreadsheetbench_verified_400.tar.gz"
)


def slugify(text: str) -> str:
    """Convert text to a URL/directory-safe slug."""
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
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


def process_spreadsheet(data_root: Path, output_base: Path) -> tuple[int, int]:
    """Process SpreadsheetBench V1: each folder has prompt.txt, *_init.xlsx, *_golden.xlsx."""
    ss_dir = data_root / "spreadsheet"
    if not ss_dir.exists():
        # The tar.gz may extract directly as "spreadsheet/" at data_root level,
        # or data_root itself might be the spreadsheet folder.
        if (data_root.parent / "spreadsheet").exists():
            ss_dir = data_root.parent / "spreadsheet"
        else:
            ss_dir = data_root

    cases = 0
    files_copied = 0
    skipped = 0

    for task_dir in sorted(ss_dir.iterdir()):
        if not task_dir.is_dir():
            continue

        task_id = task_dir.name

        # Read prompt
        prompt = read_prompt(task_dir / "prompt.txt")
        if not prompt:
            print(f"  WARNING: skipping {task_id} (no prompt.txt)")
            skipped += 1
            continue

        # Find init and golden files
        # Some folders use "*_init.xlsx", others use "initial.xlsx"
        init_files = list(task_dir.glob("*_init.xlsx")) or list(task_dir.glob("initial.xlsx"))
        # Some folders use "*_golden.xlsx", others use "golden.xlsx"
        golden_files = list(task_dir.glob("*_golden.xlsx")) or list(task_dir.glob("golden.xlsx"))
        if not init_files or not golden_files:
            print(f"  WARNING: skipping {task_id} (missing init or golden xlsx)")
            skipped += 1
            continue

        init_file = init_files[0]
        golden_file = golden_files[0]

        # Folder names are already slug-safe (e.g. "142-19", "59160")
        case_name = f"ssb1-{task_id}"
        case_dir = output_base / case_name
        case_dir.mkdir(parents=True, exist_ok=True)
        (case_dir / "input").mkdir(exist_ok=True)
        (case_dir / "output").mkdir(exist_ok=True)

        if copy_if_missing(init_file, case_dir / "input" / init_file.name):
            files_copied += 1
        if copy_if_missing(golden_file, case_dir / "output" / golden_file.name):
            files_copied += 1

        write_case_yaml(case_dir, {
            "id": case_name,
            "description": f"SpreadsheetBench V1 - {task_id}",
            "prompt": prompt,
            "input_files": [
                {"path": f"input/{init_file.name}",
                 "description": "Initial spreadsheet"},
            ],
            "output_files": [
                {"path": f"output/{golden_file.name}",
                 "description": "Golden answer spreadsheet"},
            ],
            "metadata": {
                "source": "KAKA22/SpreadsheetBench",
                "category": "spreadsheet",
                "task_id": task_id,
            },
        })
        cases += 1

    if skipped:
        print(f"  Skipped {skipped} folders with missing files")

    return cases, files_copied


def main():
    parser = argparse.ArgumentParser(
        description="Download SpreadsheetBench V1 example cases"
    )
    parser.add_argument(
        "--output", default="examples_ssbv1",
        help="Output directory for cases (default: examples_ssbv1)",
    )
    args = parser.parse_args()

    output_base = Path(args.output)

    # Download and extract
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        tar_path = tmp_path / "ssbv1.tar.gz"

        print("Downloading SpreadsheetBench V1 from HuggingFace...")
        urllib.request.urlretrieve(DATASET_URL, str(tar_path))
        print("Extracting...")

        with tarfile.open(tar_path, "r:gz") as tf:
            tf.extractall(tmp_path)

        # Locate the "spreadsheet" folder inside extracted contents
        spreadsheet_dir = tmp_path / "spreadsheet"
        if spreadsheet_dir.exists():
            data_root = tmp_path
        else:
            # May be nested inside one extra folder
            extracted_dirs = [
                d for d in tmp_path.iterdir()
                if d.is_dir() and not d.name.startswith(".")
            ]
            if not extracted_dirs:
                print("ERROR: No data found in downloaded archive")
                return
            data_root = extracted_dirs[0]

        print(f"Processing cases from: {data_root}")

        n_cases, n_files = process_spreadsheet(data_root, output_base)

    print(f"\nDone: {n_cases} cases, {n_files} files copied")
    print(f"Cases are in: {output_base.resolve()}")


if __name__ == "__main__":
    main()
