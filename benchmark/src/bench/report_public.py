from __future__ import annotations

from pathlib import Path
from typing import Iterable

import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
from pydantic import BaseModel

from .paths import PLOTS_DIR, PUBLIC_REPORT, RESULTS_DIR, RUB_RESULTS_DIR

matplotlib.use("Agg")


class MethodScore(BaseModel):
    """Aggregated benchmark scores for a method."""

    method: str
    acc_norm: float
    acc_raw: float
    acc_md: float
    md_precision: float
    avg_cost: float


class RubScore(BaseModel):
    """Aggregated RUB scores for a method."""

    method: str
    rus: float
    partial_f1: float
    avg_cost: float


class ReportData(BaseModel):
    """Combined benchmark report data for plotting."""

    core: list[MethodScore]
    rub: list[RubScore]


class ChartPaths(BaseModel):
    """Generated chart image paths for public report."""

    core_chart: Path
    markdown_chart: Path
    rub_chart: Path


def _select_methods(methods: Iterable[str]) -> list[str]:
    order = ["exstruct", "pdf", "image_vlm", "html", "openpyxl"]
    available = {m for m in methods}
    return [m for m in order if m in available]


def load_report_data() -> ReportData:
    """Load aggregated metrics from results CSV files.

    Returns:
        ReportData containing core and RUB aggregates.
    """
    core_csv = RESULTS_DIR / "results.csv"
    if not core_csv.exists():
        raise FileNotFoundError(core_csv)

    core_df = pd.read_csv(core_csv)
    core_grouped = (
        core_df.groupby("method")
        .agg(
            acc_norm=("score_norm", "mean"),
            acc_raw=("score_raw", "mean"),
            acc_md=("score_md", "mean"),
            md_precision=("score_md_precision", "mean"),
            avg_cost=("cost_usd", "mean"),
        )
        .reset_index()
    )
    core_grouped = core_grouped.fillna(0.0)

    core_methods = _select_methods(core_grouped["method"].tolist())
    core_scores = [
        MethodScore(
            method=row["method"],
            acc_norm=float(row["acc_norm"]),
            acc_raw=float(row["acc_raw"]),
            acc_md=float(row["acc_md"]),
            md_precision=float(row["md_precision"]),
            avg_cost=float(row["avg_cost"]),
        )
        for _, row in core_grouped.iterrows()
        if row["method"] in core_methods
    ]
    core_scores.sort(key=lambda m: core_methods.index(m.method))

    rub_csv = RUB_RESULTS_DIR / "rub_results.csv"
    if not rub_csv.exists():
        raise FileNotFoundError(rub_csv)

    rub_df = pd.read_csv(rub_csv)
    if "track" in rub_df.columns and (rub_df["track"] == "structure_query").any():
        rub_df = rub_df[rub_df["track"] == "structure_query"]

    rub_grouped = (
        rub_df.groupby("method")
        .agg(
            rus=("score", "mean"),
            partial_f1=("partial_f1", "mean"),
            avg_cost=("cost_usd", "mean"),
        )
        .reset_index()
    )
    rub_grouped = rub_grouped.fillna(0.0)

    rub_methods = _select_methods(rub_grouped["method"].tolist())
    rub_scores = [
        RubScore(
            method=row["method"],
            rus=float(row["rus"]),
            partial_f1=float(row["partial_f1"]),
            avg_cost=float(row["avg_cost"]),
        )
        for _, row in rub_grouped.iterrows()
        if row["method"] in rub_methods
    ]
    rub_scores.sort(key=lambda m: rub_methods.index(m.method))

    return ReportData(core=core_scores, rub=rub_scores)


def _plot_grouped_bar(
    *,
    title: str,
    ylabel: str,
    categories: list[str],
    series: dict[str, list[float]],
    out_path: Path,
) -> None:
    """Plot a grouped bar chart.

    Args:
        title: Chart title.
        ylabel: Y-axis label.
        categories: X-axis category labels.
        series: Mapping of series label to values.
        out_path: Output image path.
    """
    num_series = len(series)
    width = 0.18 if num_series > 4 else 0.22
    centers = list(range(len(categories)))

    fig, ax = plt.subplots(figsize=(9, 4.5))
    for idx, (label, values) in enumerate(series.items()):
        offset = (idx - (num_series - 1) / 2) * width
        ax.bar([c + offset for c in centers], values, width=width, label=label)

    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.set_xticks(centers)
    ax.set_xticklabels(categories, rotation=0)
    ax.set_ylim(0.0, 1.0)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    ax.legend(ncol=num_series)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(out_path, dpi=160)
    plt.close(fig)


def generate_charts(data: ReportData) -> ChartPaths:
    """Generate chart images for the public report.

    Args:
        data: Aggregated report data.

    Returns:
        ChartPaths with generated image locations.
    """
    core_chart = PLOTS_DIR / "core_benchmark.png"
    markdown_chart = PLOTS_DIR / "markdown_quality.png"
    rub_chart = PLOTS_DIR / "rub_structure_query.png"

    methods = [m.method for m in data.core]
    _plot_grouped_bar(
        title="Core Benchmark Summary",
        ylabel="Score",
        categories=methods,
        series={
            "acc_norm": [m.acc_norm for m in data.core],
            "acc_raw": [m.acc_raw for m in data.core],
            "acc_md": [m.acc_md for m in data.core],
        },
        out_path=core_chart,
    )

    _plot_grouped_bar(
        title="Markdown Evaluation Summary",
        ylabel="Score",
        categories=methods,
        series={
            "acc_md": [m.acc_md for m in data.core],
            "md_precision": [m.md_precision for m in data.core],
        },
        out_path=markdown_chart,
    )

    rub_methods = [m.method for m in data.rub]
    _plot_grouped_bar(
        title="RUB Structure Query Summary",
        ylabel="Score",
        categories=rub_methods,
        series={
            "rus": [m.rus for m in data.rub],
            "partial_f1": [m.partial_f1 for m in data.rub],
        },
        out_path=rub_chart,
    )

    return ChartPaths(
        core_chart=core_chart,
        markdown_chart=markdown_chart,
        rub_chart=rub_chart,
    )


def update_public_report(chart_paths: ChartPaths) -> Path:
    """Insert chart images into REPORT.md.

    Args:
        chart_paths: Generated chart paths.

    Returns:
        Path to updated report.
    """
    report_path = PUBLIC_REPORT
    report_text = (
        report_path.read_text(encoding="utf-8") if report_path.exists() else ""
    )

    rel_core = chart_paths.core_chart.relative_to(report_path.parent)
    rel_markdown = chart_paths.markdown_chart.relative_to(report_path.parent)
    rel_rub = chart_paths.rub_chart.relative_to(report_path.parent)

    block_lines = [
        "<!-- CHARTS_START -->",
        "## Charts",
        "",
        f"![Core Benchmark Summary]({rel_core.as_posix()})",
        f"![Markdown Evaluation Summary]({rel_markdown.as_posix()})",
        f"![RUB Structure Query Summary]({rel_rub.as_posix()})",
        "<!-- CHARTS_END -->",
        "",
    ]
    block = "\n".join(block_lines)

    if "<!-- CHARTS_START -->" in report_text and "<!-- CHARTS_END -->" in report_text:
        pre, _ = report_text.split("<!-- CHARTS_START -->", 1)
        _, post = report_text.split("<!-- CHARTS_END -->", 1)
        new_text = pre.rstrip() + "\n" + block + post.lstrip()
    else:
        new_text = report_text.rstrip() + "\n\n" + block

    report_path.write_text(new_text, encoding="utf-8")
    return report_path
