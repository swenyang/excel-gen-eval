"""JSON reporter — writes structured evaluation results."""

from __future__ import annotations

import json
from pathlib import Path

from ..models import EvalResult


def generate_json_report(
    results: list[EvalResult],
    output_path: str | Path,
) -> Path:
    """Write evaluation results as a JSON file.

    For a single result, writes the object directly.
    For multiple results, writes a list.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if len(results) == 1:
        data = _serialize(results[0])
    else:
        data = [_serialize(r) for r in results]

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False, default=str)

    return output_path


def _serialize(result: EvalResult) -> dict:
    """Convert an EvalResult to a JSON-serializable dict."""
    return json.loads(result.model_dump_json())
