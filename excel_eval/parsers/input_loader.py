"""Load grounding data from various file formats."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


def load_input_file(path: str | Path) -> str:
    """Load a single input file and return its content as text.

    Supports: CSV, JSON, JSONL, Excel (.xlsx/.xls), plain text.
    For binary/structured formats, converts to a readable text representation.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    suffix = path.suffix.lower()

    if suffix == ".csv":
        return _load_csv(path)
    elif suffix == ".json":
        return _load_json(path)
    elif suffix == ".jsonl":
        return _load_jsonl(path)
    elif suffix in (".xlsx", ".xls"):
        return _load_excel(path)
    else:
        return _load_text(path)


def load_all_inputs(file_configs: list[dict], base_dir: str | Path) -> str:
    """Load all input files and combine into a single text block.

    Each file_config should have 'path' and optionally 'description'.
    Returns a combined text with file headers.
    """
    base_dir = Path(base_dir)
    parts: list[str] = []

    for fc in file_configs:
        file_path = base_dir / fc["path"]
        desc = fc.get("description", "")
        header = f"--- Input File: {fc['path']}"
        if desc:
            header += f" ({desc})"
        header += " ---"

        try:
            content = load_input_file(file_path)
            parts.append(f"{header}\n{content}")
        except Exception as e:
            parts.append(f"{header}\n[Error loading file: {e}]")

    return "\n\n".join(parts)


def _load_csv(path: Path) -> str:
    """Load CSV file, return as CSV text."""
    return path.read_text(encoding="utf-8")


def _load_json(path: Path) -> str:
    """Load JSON file, return as pretty-printed JSON."""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return json.dumps(data, indent=2, ensure_ascii=False)


def _load_jsonl(path: Path) -> str:
    """Load JSONL file, return each line pretty-printed."""
    lines = path.read_text(encoding="utf-8").strip().split("\n")
    formatted = []
    for i, line in enumerate(lines):
        try:
            obj = json.loads(line)
            formatted.append(f"Record {i + 1}: {json.dumps(obj, ensure_ascii=False)}")
        except json.JSONDecodeError:
            formatted.append(f"Record {i + 1}: {line}")
    return "\n".join(formatted)


def _load_excel(path: Path) -> str:
    """Load Excel file as input data, return CSV-like text per sheet."""
    parts: list[str] = []
    xls = pd.ExcelFile(path)

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        csv_text = df.to_csv(index=False)
        parts.append(f"[Sheet: {sheet_name}]\n{csv_text}")

    return "\n\n".join(parts)


def _load_text(path: Path) -> str:
    """Load plain text file."""
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="latin-1")
