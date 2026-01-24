from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pydantic import BaseModel
from rich import print
from rich.console import Console
import typer

from .eval.normalize import normalize_json_text
from .eval.report import write_results_csv
from .eval.score import key_score, key_score_ordered
from .llm.openai_client import OpenAIResponsesClient
from .manifest import Case, load_manifest
from .paths import (
    DATA_DIR,
    EXTRACTED_DIR,
    PROMPTS_DIR,
    RESPONSES_DIR,
    RESULTS_DIR,
    resolve_path,
)
from .pipeline.common import ensure_dir, sha256_text, write_json
from .pipeline.exstruct_adapter import extract_exstruct
from .pipeline.html_text import html_to_text, xlsx_to_html
from .pipeline.image_render import xlsx_to_pngs_via_pdf
from .pipeline.openpyxl_pandas import extract_openpyxl
from .pipeline.pdf_text import pdf_to_text, xlsx_to_pdf

app = typer.Typer(add_completion=False)
console = Console()

METHODS_TEXT = ("exstruct", "openpyxl", "pdf", "html")
METHODS_ALL = METHODS_TEXT + ("image_vlm",)


class PromptRecord(BaseModel):
    """Prompt metadata saved for each request."""

    case_id: str
    method: str
    model: str
    temperature: float
    question: str
    prompt_hash: str
    images: list[str] | None = None


class ResponseRecord(BaseModel):
    """Response metadata saved for each request."""

    case_id: str
    method: str
    model: str
    temperature: float
    prompt_hash: str
    text: str
    input_tokens: int
    output_tokens: int
    cost_usd: float
    raw: dict[str, Any]


class ResultRow(BaseModel):
    """Evaluation row for CSV output."""

    case_id: str
    type: str
    method: str
    model: str | None
    score: float
    score_ordered: float
    ok: bool
    input_tokens: int
    output_tokens: int
    cost_usd: float
    error: str | None


def _manifest_path() -> Path:
    """Return the path to the benchmark manifest.

    Returns:
        Path to manifest.json.
    """
    return DATA_DIR / "manifest.json"


def _select_cases(manifest_cases: list[Case], case: str) -> list[Case]:
    """Select benchmark cases by id list or all.

    Args:
        manifest_cases: List of cases from the manifest.
        case: Comma-separated case ids or "all".

    Returns:
        Filtered list of cases.
    """
    if case == "all":
        return manifest_cases
    ids = {c.strip() for c in case.split(",") if c.strip()}
    return [c for c in manifest_cases if c.id in ids]


def _select_methods(method: str) -> list[str]:
    """Select methods by list or all, validating against known methods.

    Args:
        method: Comma-separated method names or "all".

    Returns:
        Ordered list of validated methods.
    """
    if method == "all":
        selected = list(METHODS_ALL)
    else:
        selected = [m.strip() for m in method.split(",") if m.strip()]

    seen: set[str] = set()
    deduped = [m for m in selected if not (m in seen or seen.add(m))]
    invalid = [m for m in deduped if m not in METHODS_ALL]
    if invalid:
        raise typer.BadParameter(
            f"Unknown method(s): {', '.join(invalid)}. Allowed: {', '.join(METHODS_ALL)}"
        )
    if not deduped:
        raise typer.BadParameter("No methods selected.")
    return deduped


def _resolve_case_path(path_str: str, *, case_id: str, label: str) -> Path | None:
    """Resolve a manifest path, warning if missing.

    Args:
        path_str: Path string from the manifest.
        case_id: Case identifier for log messages.
        label: Label for the path type (e.g., "xlsx", "truth").

    Returns:
        Resolved Path if it exists, otherwise None.
    """
    resolved = resolve_path(path_str)
    if resolved.exists():
        return resolved
    print(f"[yellow]skip: missing {label} for {case_id}: {resolved}[/yellow]")
    return None


def _reset_case_outputs(case_id: str) -> None:
    """Delete existing prompt/response logs for a case."""
    for directory in (PROMPTS_DIR, RESPONSES_DIR):
        path = directory / f"{case_id}.jsonl"
        if path.exists():
            path.unlink()


def _dump_jsonl(obj: BaseModel) -> str:
    """Serialize a record for JSONL output.

    Args:
        obj: Pydantic model to serialize.

    Returns:
        Single-line JSON string with stable key ordering.
    """
    payload = obj.model_dump(exclude_none=True)
    return json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(", ", ": ")
    )


@app.command()
def extract(case: str = "all", method: str = "all") -> None:
    """Extract contexts for selected cases and methods.

    Args:
        case: Comma-separated case ids or "all".
        method: Comma-separated method names or "all".
    """
    mf = load_manifest(_manifest_path())
    cases = _select_cases(mf.cases, case)
    if not cases:
        raise typer.BadParameter(f"No cases matched: {case}")
    methods = _select_methods(method)

    for c in cases:
        xlsx = _resolve_case_path(c.xlsx, case_id=c.id, label="xlsx")
        if not xlsx:
            continue
        console.rule(f"EXTRACT {c.id} ({xlsx.name})")

        if "exstruct" in methods:
            out_txt = EXTRACTED_DIR / "exstruct" / f"{c.id}.txt"
            extract_exstruct(xlsx, out_txt, c.sheet_scope)
            print(f"[green]exstruct -> {out_txt}[/green]")

        if "openpyxl" in methods:
            out_txt = EXTRACTED_DIR / "openpyxl" / f"{c.id}.txt"
            extract_openpyxl(xlsx, out_txt, c.sheet_scope)
            print(f"[green]openpyxl -> {out_txt}[/green]")

        if "pdf" in methods:
            out_pdf = EXTRACTED_DIR / "pdf" / f"{c.id}.pdf"
            out_txt = EXTRACTED_DIR / "pdf" / f"{c.id}.txt"
            xlsx_to_pdf(xlsx, out_pdf)
            pdf_to_text(out_pdf, out_txt)
            print(f"[green]pdf -> {out_txt}[/green]")

        if "html" in methods:
            out_html = EXTRACTED_DIR / "html" / f"{c.id}.html"
            out_txt = EXTRACTED_DIR / "html" / f"{c.id}.txt"
            xlsx_to_html(xlsx, out_html)
            html_to_text(out_html, out_txt)
            print(f"[green]html -> {out_txt}[/green]")

        if "image_vlm" in methods:
            out_dir = EXTRACTED_DIR / "image_vlm" / c.id
            pngs = xlsx_to_pngs_via_pdf(
                xlsx, out_dir, dpi=c.render.dpi, max_pages=c.render.max_pages
            )
            write_json(out_dir / "images.json", {"images": [str(p) for p in pngs]})
            print(f"[green]image_vlm -> {len(pngs)} png(s) in {out_dir}[/green]")


@app.command()
def ask(
    case: str = "all",
    method: str = "all",
    model: str = "gpt-4o",
    temperature: float = 0.0,
) -> None:
    """Run LLM extraction against prepared contexts.

    Args:
        case: Comma-separated case ids or "all".
        method: Comma-separated method names or "all".
        model: OpenAI model name.
        temperature: Sampling temperature for the model.
    """
    mf = load_manifest(_manifest_path())
    cases = _select_cases(mf.cases, case)
    if not cases:
        raise typer.BadParameter(f"No cases matched: {case}")
    methods = _select_methods(method)

    client = OpenAIResponsesClient()
    ensure_dir(PROMPTS_DIR)
    ensure_dir(RESPONSES_DIR)

    for c in cases:
        console.rule(f"ASK {c.id}")
        q = c.question
        _reset_case_outputs(c.id)

        for m in methods:
            if m == "image_vlm":
                img_dir = EXTRACTED_DIR / "image_vlm" / c.id
                images_json = img_dir / "images.json"
                if not images_json.exists():
                    print(f"[yellow]skip: missing images for {c.id}[/yellow]")
                    continue
                imgs = json.loads(images_json.read_text(encoding="utf-8"))["images"]
                img_paths = [Path(p) for p in imgs]
                if not img_paths:
                    print(f"[yellow]skip: no images for {c.id}[/yellow]")
                    continue
                prompt_hash = sha256_text(
                    q + "|" + "|".join([p.name for p in img_paths])
                )
                prompt_rec = PromptRecord(
                    case_id=c.id,
                    method=m,
                    model=model,
                    temperature=temperature,
                    question=q,
                    prompt_hash=prompt_hash,
                    images=[p.name for p in img_paths],
                )
                res = client.ask_images(
                    model=model,
                    question=q,
                    image_paths=img_paths,
                    temperature=temperature,
                )
            else:
                txt_path = EXTRACTED_DIR / m / f"{c.id}.txt"
                if not txt_path.exists():
                    print(f"[yellow]skip: missing context for {c.id} ({m})[/yellow]")
                    continue
                context = txt_path.read_text(encoding="utf-8")
                prompt_hash = sha256_text(q + "|" + context)
                prompt_rec = PromptRecord(
                    case_id=c.id,
                    method=m,
                    model=model,
                    temperature=temperature,
                    question=q,
                    prompt_hash=prompt_hash,
                )
                res = client.ask_text(
                    model=model,
                    question=q,
                    context_text=context,
                    temperature=temperature,
                )

            prompt_file = PROMPTS_DIR / f"{c.id}.jsonl"
            resp_file = RESPONSES_DIR / f"{c.id}.jsonl"
            resp_rec = ResponseRecord(
                case_id=c.id,
                method=m,
                model=model,
                temperature=temperature,
                prompt_hash=prompt_hash,
                text=res.text,
                input_tokens=res.input_tokens,
                output_tokens=res.output_tokens,
                cost_usd=res.cost_usd,
                raw=res.raw,
            )

            prompt_line = _dump_jsonl(prompt_rec)
            resp_line = _dump_jsonl(resp_rec)
            with prompt_file.open("a", encoding="utf-8") as f:
                f.write(prompt_line + "\n")
            with resp_file.open("a", encoding="utf-8") as f:
                f.write(resp_line + "\n")

            print(
                f"[cyan]{c.id} {m}[/cyan] tokens(in/out)={res.input_tokens}/{res.output_tokens} cost=${res.cost_usd:.6f}"
            )


@app.command()
def eval(case: str = "all", method: str = "all") -> None:
    """Evaluate the latest responses and write results CSV.

    Args:
        case: Comma-separated case ids or "all".
        method: Comma-separated method names or "all".
    """
    mf = load_manifest(_manifest_path())
    cases = _select_cases(mf.cases, case)
    if not cases:
        raise typer.BadParameter(f"No cases matched: {case}")
    methods = _select_methods(method)

    rows: list[ResultRow] = []

    for c in cases:
        truth_path = _resolve_case_path(c.truth, case_id=c.id, label="truth")
        if not truth_path:
            continue
        truth = json.loads(truth_path.read_text(encoding="utf-8"))
        resp_file = RESPONSES_DIR / f"{c.id}.jsonl"
        if not resp_file.exists():
            print(f"[yellow]skip: no responses for {c.id}[/yellow]")
            continue

        latest: dict[str, dict[str, Any]] = {}
        for line in resp_file.read_text(encoding="utf-8").splitlines():
            rec = json.loads(line)
            if rec.get("method") in methods:
                latest[rec["method"]] = rec

        for m, rec in latest.items():
            ok = False
            score = 0.0
            score_ordered = 0.0
            err: str | None = None
            try:
                pred_obj = normalize_json_text(rec["text"])
                score = key_score(truth, pred_obj)
                score_ordered = key_score_ordered(truth, pred_obj)
                ok = score == 1.0
            except Exception as exc:
                err = str(exc)

            rows.append(
                ResultRow(
                    case_id=c.id,
                    type=c.type,
                    method=m,
                    model=rec.get("model"),
                    score=score,
                    score_ordered=score_ordered,
                    ok=ok,
                    input_tokens=int(rec.get("input_tokens", 0)),
                    output_tokens=int(rec.get("output_tokens", 0)),
                    cost_usd=float(rec.get("cost_usd", 0.0)),
                    error=err,
                )
            )

    out_csv = RESULTS_DIR / "results.csv"
    write_results_csv([row.model_dump() for row in rows], out_csv)
    print(f"[green]Wrote {out_csv} ({len(rows)} rows)[/green]")


@app.command()
def report() -> None:
    """Generate a Markdown report from the results CSV."""
    csv_path = RESULTS_DIR / "results.csv"
    if not csv_path.exists():
        raise typer.Exit(code=1)

    import pandas as pd

    df = pd.read_csv(csv_path)
    score_col = "score" if "score" in df.columns else "ok"
    agg: dict[str, tuple[str, str]] = {
        "acc": (score_col, "mean"),
        "avg_in": ("input_tokens", "mean"),
        "avg_cost": ("cost_usd", "mean"),
        "n": (score_col, "count"),
    }
    if "score_ordered" in df.columns:
        agg["acc_ordered"] = ("score_ordered", "mean")
    g = df.groupby("method").agg(**agg).reset_index()

    md_lines = []
    md_lines.append("# Benchmark Report")
    md_lines.append("")
    md_lines.append("## Summary by method")
    md_lines.append("")
    md_lines.append(g.to_markdown(index=False))
    md_lines.append("")
    out_md = RESULTS_DIR / "report.md"
    out_md.write_text("\n".join(md_lines), encoding="utf-8")
    print(f"[green]Wrote {out_md}[/green]")
    print("[cyan]Summary (from report.md)[/cyan]")
    print(g.to_markdown(index=False))


if __name__ == "__main__":
    app()
