from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]  # benchmark/
DATA_DIR = ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
TRUTH_DIR = DATA_DIR / "truth"

OUT_DIR = ROOT / "outputs"
EXTRACTED_DIR = OUT_DIR / "extracted"
PROMPTS_DIR = OUT_DIR / "prompts"
RESPONSES_DIR = OUT_DIR / "responses"
MARKDOWN_DIR = OUT_DIR / "markdown"
MARKDOWN_RESPONSES_DIR = MARKDOWN_DIR / "responses"
MARKDOWN_FULL_DIR = OUT_DIR / "markdown_full"
MARKDOWN_FULL_RESPONSES_DIR = MARKDOWN_FULL_DIR / "responses"
RESULTS_DIR = OUT_DIR / "results"
PLOTS_DIR = OUT_DIR / "plots"
PUBLIC_REPORT = ROOT / "REPORT.md"
RUB_DIR = ROOT / "rub"
RUB_MANIFEST = RUB_DIR / "manifest.json"
RUB_TRUTH_DIR = RUB_DIR / "truth"
RUB_SCHEMA_DIR = RUB_DIR / "schemas"
RUB_OUT_DIR = OUT_DIR / "rub"
RUB_PROMPTS_DIR = RUB_OUT_DIR / "prompts"
RUB_RESPONSES_DIR = RUB_OUT_DIR / "responses"
RUB_RESULTS_DIR = RUB_OUT_DIR / "results"


def resolve_path(path: str | Path) -> Path:
    """Resolve a path relative to the benchmark root when needed.

    Args:
        path: Path string or Path instance from the manifest.

    Returns:
        Resolved Path anchored to the benchmark root when relative.
    """
    candidate = Path(path)
    if candidate.is_absolute():
        return candidate
    return ROOT / candidate
