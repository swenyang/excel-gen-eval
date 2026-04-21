"""Regression test runner — validates eval scores against expected ranges."""

import asyncio
import sys
from pathlib import Path

import yaml

from excel_eval.config import load_global_config
from excel_eval.pipeline import Pipeline


def main():
    config_path = Path("eval_config.yaml")
    if not config_path.exists():
        print("ERROR: eval_config.yaml not found")
        sys.exit(1)

    regression_path = Path("tests/regression/regression_cases.yaml")
    with open(regression_path, encoding="utf-8") as f:
        spec = yaml.safe_load(f)

    # Check if test data exists
    missing = []
    for case_spec in spec["cases"]:
        case_dir = Path(case_spec["case_dir"])
        if not case_dir.exists():
            missing.append(str(case_dir))
        else:
            xlsx_files = list(case_dir.rglob("*.xlsx"))
            if not xlsx_files:
                missing.append(f"{case_dir} (no .xlsx files)")

    if missing:
        print("WARNING: Some test cases are missing data files:")
        for m in missing:
            print(f"  {m}")
        print("\nTo download test data, run:")
        print("  python scripts/download_gdpval.py    # GDPVal cases")
        print("  python scripts/download_ssbv2.py --output examples_ssbv2  # SSBv2 cases")
        print()

    config = load_global_config(str(config_path))
    pipeline = Pipeline(config)

    failures = []

    async def run():
        for case_spec in spec["cases"]:
            case_dir = case_spec["case_dir"]
            desc = case_spec.get("description", "")
            expected = case_spec["expected"]

            print(f"\nRunning: {case_dir} ({desc})")
            result = await pipeline.evaluate(case_dir)

            # Check each dimension
            for dim, exp_range in expected.items():
                if dim == "overall_min":
                    actual = result.summary.overall_weighted_avg or 0
                    if actual < exp_range:
                        msg = f"  FAIL: overall {actual:.2f} < expected min {exp_range}"
                        print(msg)
                        failures.append(msg)
                    else:
                        print(f"  OK: overall {actual:.2f} >= {exp_range}")
                    continue

                dr = result.dimensions.get(dim)
                if exp_range is None:
                    if dr and dr.score is not None:
                        msg = f"  FAIL: {dim} = {dr.score} (expected N/A)"
                        print(msg)
                        failures.append(msg)
                    else:
                        print(f"  OK: {dim} = N/A")
                    continue

                actual = dr.score if dr else None
                if actual is None:
                    msg = f"  FAIL: {dim} = N/A (expected {exp_range})"
                    print(msg)
                    failures.append(msg)
                elif actual < exp_range[0] or actual > exp_range[1]:
                    msg = f"  FAIL: {dim} = {actual} (expected {exp_range[0]}-{exp_range[1]})"
                    print(msg)
                    failures.append(msg)
                else:
                    print(f"  OK: {dim} = {actual} (in range {exp_range})")

    asyncio.run(run())

    print(f"\n{'='*40}")
    if failures:
        print(f"REGRESSION DETECTED: {len(failures)} failure(s)")
        for f in failures:
            print(f"  {f}")
        sys.exit(1)
    else:
        print("ALL REGRESSION TESTS PASSED")


if __name__ == "__main__":
    main()
