"""CLI entry point for excel-eval."""

from __future__ import annotations

import asyncio
import logging
import sys
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table

from .config import load_global_config, discover_cases
from .models import DimensionName, EvalResult, GlobalConfig
from .pipeline import Pipeline
from .reporters.json_reporter import generate_json_report
from .reporters.excel_reporter import generate_excel_report

console = Console()


def _setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)-8s %(name)s — %(message)s",
        datefmt="%H:%M:%S",
    )


def _print_result(result: EvalResult) -> None:
    """Print a single evaluation result to the console."""
    console.print(f"\n[bold]Case:[/bold] {result.case_id}")
    console.print(f"[bold]Scenario:[/bold] {result.scenario.detected.value} "
                  f"(confidence: {result.scenario.confidence:.0%})")

    table = Table(title="Dimension Scores", show_lines=True)
    table.add_column("Dimension", style="bold")
    table.add_column("Score", justify="center")
    table.add_column("Status", justify="center")

    for dim in DimensionName:
        dr = result.dimensions.get(dim.value)
        if dr is None:
            continue

        score_str = str(dr.score) if dr.score is not None else "N/A"
        if dr.score is not None:
            colors = {0: "red", 1: "red", 2: "yellow", 3: "green", 4: "bold green"}
            color = colors.get(dr.score, "white")
            score_str = f"[{color}]{dr.score}/4[/{color}]"

        status_color = {
            "success": "green", "error": "red", "na": "dim", "skipped": "dim"
        }.get(dr.status.value, "white")

        table.add_row(
            dim.value.replace("_", " ").title(),
            score_str,
            f"[{status_color}]{dr.status.value}[/{status_color}]",
        )

    console.print(table)

    # Summary
    s = result.summary
    console.print(f"\n[bold]Data & Content Avg:[/bold] {s.data_content_avg:.2f}" if s.data_content_avg else "")
    console.print(f"[bold]Structure & Usability Avg:[/bold] {s.structure_usability_avg:.2f}" if s.structure_usability_avg else "")
    console.print(f"[bold green]Overall Weighted Avg:[/bold green] {s.overall_weighted_avg:.2f}" if s.overall_weighted_avg else "")
    console.print(f"[dim]Cost: ${result.cost.total_cost_estimate:.4f} "
                  f"({result.cost.total_input_tokens + result.cost.total_output_tokens} tokens)[/dim]")

    # Performance info
    perf = result.metadata.get("performance")
    if perf:
        s1 = perf.get("stage1_ms", 0) / 1000
        total = perf.get("total_wall_ms", 0) / 1000
        screenshots = perf.get("screenshots_count", 0)
        console.print(f"[dim]Perf: stage1={s1:.1f}s, total={total:.1f}s, screenshots={screenshots}[/dim]")


def _generate_reports(results: list[EvalResult], output_dir: Path, formats: list[str]) -> None:
    """Generate reports in requested formats."""
    output_dir.mkdir(parents=True, exist_ok=True)


def _print_batch_summary(results: list[EvalResult]) -> None:
    """Print a summary table for batch evaluation results."""
    console.print("\n[bold]═══ Batch Summary ═══[/bold]")

    table = Table(title="Per-Case Performance", show_lines=False)
    table.add_column("Case", style="bold", max_width=35)
    table.add_column("Score", justify="right")
    table.add_column("Tokens", justify="right")
    table.add_column("Cost", justify="right")
    table.add_column("Stage1", justify="right")
    table.add_column("Total", justify="right")

    total_tokens = 0
    total_cost = 0.0
    total_wall = 0
    scores = []

    for r in sorted(results, key=lambda x: x.summary.overall_weighted_avg or 0):
        ow = r.summary.overall_weighted_avg
        tokens = r.cost.total_input_tokens + r.cost.total_output_tokens
        cost = r.cost.total_cost_estimate
        perf = r.metadata.get("performance", {})
        s1 = perf.get("stage1_ms", 0)
        wall = perf.get("total_wall_ms", 0)

        total_tokens += tokens
        total_cost += cost
        total_wall += wall
        if ow is not None:
            scores.append(ow)

        score_str = f"{ow:.2f}" if ow is not None else "N/A"
        table.add_row(
            r.case_id,
            score_str,
            f"{tokens:,}",
            f"${cost:.3f}",
            f"{s1/1000:.1f}s",
            f"{wall/1000:.1f}s",
        )

    console.print(table)

    # Totals
    mean_score = sum(scores) / len(scores) if scores else 0
    console.print(f"\n[bold]Mean score:[/bold] {mean_score:.2f}")
    console.print(f"[bold]Total tokens:[/bold] {total_tokens:,}")
    console.print(f"[bold]Total cost:[/bold] ${total_cost:.2f}")
    console.print(f"[bold]Total wall time:[/bold] {total_wall/1000:.0f}s ({total_wall/60000:.1f}min)")


    if "json" in formats:
        path = generate_json_report(results, output_dir / "eval_result.json")
        console.print(f"  [bold]JSON:[/bold] {path}")

    if "html" in formats:
        try:
            from .reporters.html_reporter import generate_html_report
            path = generate_html_report(results, output_dir / "eval_report.html")
            console.print(f"  [bold]HTML:[/bold] {path}")
        except ImportError:
            console.print("  HTML reporter not available", style="yellow")

    if "excel" in formats:
        path = generate_excel_report(results, output_dir / "eval_report.xlsx")
        console.print(f"  [bold]Excel:[/bold] {path}")


# ── CLI Commands ───────────────────────────────────────────────────────────


@click.group()
@click.version_option(version="0.1.0")
def main():
    """Excel Gen Eval — Evaluate AI-generated Excel files."""
    pass


@main.command()
@click.argument("path", type=click.Path(exists=True))
@click.option("--batch", is_flag=True, help="Evaluate all cases under PATH")
@click.option("--config", "config_path", type=click.Path(), help="Global config YAML")
@click.option("--dimensions", "dims", help="Comma-separated dimension filter")
@click.option("--output", "output_dir", default="./results", help="Output directory")
@click.option("--format", "formats", default="json,html,excel", help="Output formats (comma-separated)")
@click.option("--runs", default=1, type=int, help="Eval runs per case (>1 takes median; use for stability diagnostics, not routine)")
@click.option("--parallel", default=1, type=int, help="Number of cases to evaluate concurrently in batch mode (default: 1)")
@click.option("-v", "--verbose", is_flag=True, help="Verbose logging")
def run(path, batch, config_path, dims, output_dir, formats, runs, parallel, verbose):
    """Evaluate Excel file(s) in a test case directory."""
    _setup_logging(verbose)

    # Auto-discover config: explicit > eval_config.yaml in cwd > defaults
    if not config_path:
        auto_config = Path("eval_config.yaml")
        if auto_config.exists():
            config_path = str(auto_config)

    config = load_global_config(config_path) if config_path else GlobalConfig()

    format_list = [f.strip() for f in formats.split(",")]
    output_path = Path(output_dir)

    async def _run():
        pipeline = Pipeline(config)

        if batch:
            console.print(f"[bold]Batch evaluation:[/bold] {path} (parallel={parallel})")
            results = await pipeline.evaluate_batch(
                path, num_runs=runs, parallel=parallel, output_dir=output_path,
            )
        else:
            console.print(f"[bold]Evaluating:[/bold] {path}")
            case_debug = output_path / "debug"
            results = [await pipeline.evaluate(path, num_runs=runs, output_dir=case_debug)]

        for r in results:
            _print_result(r)

        # Batch summary with performance stats
        if batch and len(results) > 1:
            _print_batch_summary(results)

        console.print("\n[bold]Generating reports:[/bold]")
        _generate_reports(results, output_path, format_list)

    asyncio.run(_run())


@main.command("dimensions")
def list_dimensions():
    """List all evaluation dimensions."""
    table = Table(title="Evaluation Dimensions")
    table.add_column("Block", style="bold")
    table.add_column("Dimension")
    table.add_column("Key")

    data_content = [
        ("Data Accuracy", "data_accuracy"),
        ("Completeness", "completeness"),
        ("Formula & Logic", "formula_logic"),
        ("Relevance", "relevance"),
    ]
    structure = [
        ("Sheet Organization", "sheet_organization"),
        ("Table Structure", "table_structure"),
        ("Chart Appropriateness", "chart_appropriateness"),
        ("Professional Formatting", "professional_formatting"),
    ]

    for name, key in data_content:
        table.add_row("Data & Content", name, key)
    for name, key in structure:
        table.add_row("Structure & Usability", name, key)

    console.print(table)


if __name__ == "__main__":
    main()
