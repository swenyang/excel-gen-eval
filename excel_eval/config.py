"""Configuration loading and validation."""

from __future__ import annotations

import os
from pathlib import Path

import yaml

from .models import CaseConfig, GlobalConfig, LLMConfig


def load_global_config(path: str | Path | None = None) -> GlobalConfig:
    """Load global configuration from a YAML file.

    If no path is provided, returns default configuration.
    The LLM API key is resolved from environment variable if api_key_env is set.
    """
    if path is None:
        return GlobalConfig()

    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    with open(path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    config = GlobalConfig(**raw)
    _resolve_api_key(config.llm)
    return config


def load_case_config(case_dir: str | Path) -> CaseConfig:
    """Load a test case configuration from case.yaml in the given directory.

    Resolves relative paths in input_files and output_files against the case directory.
    """
    case_dir = Path(case_dir)
    yaml_path = case_dir / "case.yaml"

    if not yaml_path.exists():
        raise FileNotFoundError(f"case.yaml not found in {case_dir}")

    with open(yaml_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    return CaseConfig(**raw)


def discover_cases(root_dir: str | Path) -> list[Path]:
    """Discover all test case directories under root_dir.

    A directory is considered a test case if it contains a case.yaml file.
    Returns sorted list of case directories.
    """
    root = Path(root_dir)
    cases = sorted(
        p.parent for p in root.rglob("case.yaml")
    )
    return cases


def _resolve_api_key(llm_config: LLMConfig) -> None:
    """Resolve API key from environment variable if configured."""
    if llm_config.api_key:
        return
    if llm_config.api_key_env:
        key = os.environ.get(llm_config.api_key_env)
        if key:
            llm_config.api_key = key
        else:
            raise ValueError(
                f"Environment variable '{llm_config.api_key_env}' not set"
            )
