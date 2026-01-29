from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pydantic import BaseModel
from rich import print
from rich.console import Console
import typer

from .eval.markdown_render import render_markdown
from .eval.markdown_score import markdown_coverage_score, markdown_precision_score
from .eval.normalize import normalize_json_text
from .eval.normalization_rules import load_ruleset
from .eval.report import write_results_csv
from .eval.raw_match import raw_coverage_score, raw_precision_score
from .eval.score import (
    key_score,
    key_score_normalized,
    key_score_ordered,
    key_score_ordered_normalized,
)
from .llm.openai_client import OpenAIResponsesClient
from .manifest import Case, load_manifest
from .paths import (
    DATA_DIR,
    EXTRACTED_DIR,
    MARKDOWN_DIR,
    MARKDOWN_FULL_DIR,
    MARKDOWN_FULL_RESPONSES_DIR,
    MARKDOWN_RESPONSES_DIR,
    PROMPTS_DIR,
    RESPONSES_DIR,
    RESULTS_DIR,
    RUB_MANIFEST,
    RUB_OUT_DIR,
    RUB_PROMPTS_DIR,
    RUB_RESPONSES_DIR,
    RUB_RESULTS_DIR,
    resolve_path,
)
from .pipeline.common import ensure_dir, sha256_text, write_json
from .pipeline.exstruct_adapter import extract_exstruct
from .pipeline.html_text import html_to_text, xlsx_to_html
from .pipeline.image_render import xlsx_to_pngs_via_pdf
from .pipeline.openpyxl_pandas import extract_openpyxl
from .pipeline.pdf_text import pdf_to_text, xlsx_to_pdf
from .report_public import generate_charts, load_report_data, update_public_report
from .rub.manifest import RubTask, load_rub_manifest
from .rub.score import RubPartialScore, score_exact, score_partial

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


class MarkdownRecord(BaseModel):
    """Markdown conversion metadata saved for each request."""

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


class RubResponseRecord(BaseModel):
    """RUB response metadata saved for each request."""

    task_id: str
    source_case_id: str
    method: str
    model: str
    temperature: float
    prompt_hash: str
    question: str
    text: str
    input_tokens: int
    output_tokens: int
    cost_usd: float
    raw: dict[str, Any]


class RubResultRow(BaseModel):
    """RUB evaluation row for CSV output."""

    task_id: str
    source_case_id: str
    type: str
    track: str
    method: str
    model: str | None
    score: float
    partial_precision: float | None = None
    partial_recall: float | None = None
    partial_f1: float | None = None
    ok: bool
    input_tokens: int
    output_tokens: int
    cost_usd: float
    error: str | None


class ResultRow(BaseModel):
    """Evaluation row for CSV output."""

    case_id: str
    type: str
    method: str
    model: str | None
    score: float
    score_ordered: float
    score_norm: float | None = None
    score_norm_ordered: float | None = None
    score_raw: float | None = None
    score_raw_precision: float | None = None
    score_md: float | None = None
    score_md_precision: float | None = None
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


def _rub_manifest_path(manifest_path: str | None) -> Path:
    """Return the path to the RUB manifest.

    Args:
        manifest_path: Optional override path from CLI.

    Returns:
        Path to the RUB manifest file.
    """
    if manifest_path:
        return resolve_path(manifest_path)
    return RUB_MANIFEST


def _select_tasks(tasks: list[RubTask], task: str) -> list[RubTask]:
    """Select RUB tasks by id list or all.

    Args:
        tasks: Task list from the RUB manifest.
        task: Comma-separated task ids or "all".

    Returns:
        Filtered list of tasks.
    """
    if task == "all":
        return tasks
    ids = {t.strip() for t in task.split(",") if t.strip()}
    return [t for t in tasks if t.id in ids]


def _resolve_task_path(path_str: str, *, task_id: str, label: str) -> Path | None:
    """Resolve a RUB manifest path, warning if missing.

    Args:
        path_str: Path string from the manifest.
        task_id: Task identifier for log messages.
        label: Label for the path type (e.g., "truth").

    Returns:
        Resolved Path if it exists, otherwise None.
    """
    resolved = resolve_path(path_str)
    if resolved.exists():
        return resolved
    print(f"[yellow]skip: missing {label} for {task_id}: {resolved}[/yellow]")
    return None


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


def _reset_rub_outputs(task_id: str) -> None:
    """Delete existing RUB prompt/response logs for a task."""
    for directory in (RUB_PROMPTS_DIR, RUB_RESPONSES_DIR):
        path = directory / f"{task_id}.jsonl"
        if path.exists():
            path.unlink()


def _reset_markdown_outputs(case_id: str) -> None:
    """Delete existing markdown logs for a case."""
    path = MARKDOWN_RESPONSES_DIR / f"{case_id}.jsonl"
    if path.exists():
        path.unlink()


def _reset_markdown_full_outputs(case_id: str) -> None:
    """Delete existing full-markdown logs for a case."""
    path = MARKDOWN_FULL_RESPONSES_DIR / f"{case_id}.jsonl"
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
    total_cost = 0.0
    total_calls = 0

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

            total_cost += res.cost_usd
            total_calls += 1
            print(
                f"[cyan]{c.id} {m}[/cyan] tokens(in/out)={res.input_tokens}/{res.output_tokens} cost=${res.cost_usd:.6f}"
            )
    print(f"[green]Total cost: ${total_cost:.6f} ({total_calls} call(s))[/green]")


@app.command()
def markdown(
    case: str = "all",
    method: str = "all",
    model: str = "gpt-4o",
    temperature: float = 0.0,
    use_llm: bool = True,
) -> None:
    """Generate Markdown outputs from the latest JSON responses.

    Args:
        case: Comma-separated case ids or "all".
        method: Comma-separated method names or "all".
        model: OpenAI model name for Markdown conversion.
        temperature: Sampling temperature for the model.
        use_llm: If True, call the LLM for conversion; otherwise use renderer.
    """
    mf = load_manifest(_manifest_path())
    cases = _select_cases(mf.cases, case)
    if not cases:
        raise typer.BadParameter(f"No cases matched: {case}")
    methods = _select_methods(method)

    client = OpenAIResponsesClient()
    ensure_dir(MARKDOWN_DIR)
    ensure_dir(MARKDOWN_RESPONSES_DIR)
    total_cost = 0.0
    total_calls = 0

    for c in cases:
        console.rule(f"MARKDOWN {c.id}")
        resp_file = RESPONSES_DIR / f"{c.id}.jsonl"
        if not resp_file.exists():
            print(f"[yellow]skip: no responses for {c.id}[/yellow]")
            continue
        _reset_markdown_outputs(c.id)
        latest: dict[str, dict[str, Any]] = {}
        for line in resp_file.read_text(encoding="utf-8").splitlines():
            rec = json.loads(line)
            if rec.get("method") in methods:
                latest[rec["method"]] = rec

        case_dir = MARKDOWN_DIR / c.id
        ensure_dir(case_dir)
        md_file = MARKDOWN_RESPONSES_DIR / f"{c.id}.jsonl"

        for m, rec in latest.items():
            try:
                pred_obj = normalize_json_text(rec["text"])
                json_text = json.dumps(pred_obj, ensure_ascii=False)
                prompt_hash = sha256_text(json_text)
                if use_llm:
                    if client is None:
                        raise RuntimeError(
                            "LLM client unavailable for markdown conversion."
                        )
                    res = client.ask_markdown(
                        model=model, json_text=json_text, temperature=temperature
                    )
                    md_text = res.text
                    md_rec = MarkdownRecord(
                        case_id=c.id,
                        method=m,
                        model=model,
                        temperature=temperature,
                        prompt_hash=prompt_hash,
                        text=md_text,
                        input_tokens=res.input_tokens,
                        output_tokens=res.output_tokens,
                        cost_usd=res.cost_usd,
                        raw=res.raw,
                    )
                    total_cost += res.cost_usd
                    total_calls += 1
                    line = _dump_jsonl(md_rec)
                    with md_file.open("a", encoding="utf-8") as f:
                        f.write(line + "\n")
                else:
                    md_text = render_markdown(pred_obj, title=c.id)

                out_md = case_dir / f"{m}.md"
                out_md.write_text(md_text, encoding="utf-8")
                print(f"[green]{c.id} {m} -> {out_md}[/green]")
            except Exception as exc:
                print(f"[yellow]skip: markdown {c.id} {m} ({exc})[/yellow]")

    if use_llm:
        print(
            f"[green]Markdown cost: ${total_cost:.6f} ({total_calls} call(s))[/green]"
        )


@app.command()
def markdown_full(
    case: str = "all",
    method: str = "all",
    model: str = "gpt-4o",
    temperature: float = 0.0,
) -> None:
    """Generate full-document Markdown from extracted contexts.

    Args:
        case: Comma-separated case ids or "all".
        method: Comma-separated method names or "all".
        model: OpenAI model name for Markdown conversion.
        temperature: Sampling temperature for the model.
    """
    mf = load_manifest(_manifest_path())
    cases = _select_cases(mf.cases, case)
    if not cases:
        raise typer.BadParameter(f"No cases matched: {case}")
    methods = _select_methods(method)

    client = OpenAIResponsesClient()
    ensure_dir(MARKDOWN_FULL_DIR)
    ensure_dir(MARKDOWN_FULL_RESPONSES_DIR)
    total_cost = 0.0
    total_calls = 0

    for c in cases:
        console.rule(f"MARKDOWN FULL {c.id}")
        _reset_markdown_full_outputs(c.id)
        case_dir = MARKDOWN_FULL_DIR / c.id
        ensure_dir(case_dir)
        md_file = MARKDOWN_FULL_RESPONSES_DIR / f"{c.id}.jsonl"

        for m in methods:
            try:
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
                    prompt_hash = sha256_text("|".join([p.name for p in img_paths]))
                    res = client.ask_markdown_images(
                        model=model, image_paths=img_paths, temperature=temperature
                    )
                else:
                    txt_path = EXTRACTED_DIR / m / f"{c.id}.txt"
                    if not txt_path.exists():
                        print(
                            f"[yellow]skip: missing context for {c.id} ({m})[/yellow]"
                        )
                        continue
                    context_text = txt_path.read_text(encoding="utf-8")
                    prompt_hash = sha256_text(context_text)
                    res = client.ask_markdown_from_text(
                        model=model,
                        context_text=context_text,
                        temperature=temperature,
                    )

                md_text = res.text
                md_rec = MarkdownRecord(
                    case_id=c.id,
                    method=m,
                    model=model,
                    temperature=temperature,
                    prompt_hash=prompt_hash,
                    text=md_text,
                    input_tokens=res.input_tokens,
                    output_tokens=res.output_tokens,
                    cost_usd=res.cost_usd,
                    raw=res.raw,
                )
                total_cost += res.cost_usd
                total_calls += 1
                line = _dump_jsonl(md_rec)
                with md_file.open("a", encoding="utf-8") as f:
                    f.write(line + "\n")

                out_md = case_dir / f"{m}.md"
                out_md.write_text(md_text, encoding="utf-8")
                print(f"[green]{c.id} {m} -> {out_md}[/green]")
            except Exception as exc:
                print(f"[yellow]skip: markdown full {c.id} {m} ({exc})[/yellow]")

    print(
        f"[green]Markdown full cost: ${total_cost:.6f} ({total_calls} call(s))[/green]"
    )


@app.command()
def rub_ask(
    task: str = "all",
    method: str = "all",
    model: str = "gpt-4o",
    temperature: float = 0.0,
    context: str = "partial",
    manifest: str | None = None,
) -> None:
    """Run RUB Stage B queries using Markdown outputs as context.

    Args:
        task: Comma-separated task ids or "all".
        method: Comma-separated method names or "all".
        model: OpenAI model name for Stage B queries.
        temperature: Sampling temperature for the model.
        context: Markdown source ("partial" or "full").
        manifest: Optional RUB manifest path override.
    """
    rub_manifest = load_rub_manifest(_rub_manifest_path(manifest))
    tasks = _select_tasks(rub_manifest.tasks, task)
    if not tasks:
        raise typer.BadParameter(f"No tasks matched: {task}")
    methods = _select_methods(method)
    context_key = context.lower().strip()
    if context_key not in {"partial", "full"}:
        raise typer.BadParameter(f"Invalid context: {context}")
    md_root = MARKDOWN_DIR if context_key == "partial" else MARKDOWN_FULL_DIR

    ensure_dir(RUB_OUT_DIR)
    ensure_dir(RUB_PROMPTS_DIR)
    ensure_dir(RUB_RESPONSES_DIR)

    client = OpenAIResponsesClient()
    total_cost = 0.0
    total_calls = 0

    for t in tasks:
        console.rule(f"RUB {t.id}")
        _reset_rub_outputs(t.id)
        resp_file = RUB_RESPONSES_DIR / f"{t.id}.jsonl"
        for m in methods:
            md_path = md_root / t.source_case_id / f"{m}.md"
            if not md_path.exists():
                print(f"[yellow]skip: missing markdown {t.id} {m}[/yellow]")
                continue
            context_text = md_path.read_text(encoding="utf-8")
            prompt_hash = sha256_text(f"{t.question}\n{context_text}")
            try:
                res = client.ask_text(
                    model=model,
                    question=t.question,
                    context_text=context_text,
                    temperature=temperature,
                )
                rec = RubResponseRecord(
                    task_id=t.id,
                    source_case_id=t.source_case_id,
                    method=m,
                    model=model,
                    temperature=temperature,
                    prompt_hash=prompt_hash,
                    question=t.question,
                    text=res.text,
                    input_tokens=res.input_tokens,
                    output_tokens=res.output_tokens,
                    cost_usd=res.cost_usd,
                    raw=res.raw,
                )
                line = _dump_jsonl(rec)
                with resp_file.open("a", encoding="utf-8") as f:
                    f.write(line + "\n")
                total_cost += res.cost_usd
                total_calls += 1
                print(f"[green]{t.id} {m} -> {resp_file}[/green]")
            except Exception as exc:
                print(f"[yellow]skip: rub {t.id} {m} ({exc})[/yellow]")

    print(f"[green]RUB cost: ${total_cost:.6f} ({total_calls} call(s))[/green]")


@app.command()
def rub_eval(
    task: str = "all", method: str = "all", manifest: str | None = None
) -> None:
    """Evaluate RUB responses and write results CSV.

    Args:
        task: Comma-separated task ids or "all".
        method: Comma-separated method names or "all".
        manifest: Optional RUB manifest path override.
    """
    rub_manifest = load_rub_manifest(_rub_manifest_path(manifest))
    tasks = _select_tasks(rub_manifest.tasks, task)
    if not tasks:
        raise typer.BadParameter(f"No tasks matched: {task}")
    methods = _select_methods(method)

    rows: list[RubResultRow] = []
    for t in tasks:
        truth_path = _resolve_task_path(t.truth, task_id=t.id, label="truth")
        if not truth_path:
            continue
        truth = json.loads(truth_path.read_text(encoding="utf-8"))

        resp_file = RUB_RESPONSES_DIR / f"{t.id}.jsonl"
        if not resp_file.exists():
            print(f"[yellow]skip: no RUB responses for {t.id}[/yellow]")
            continue
        latest: dict[str, dict[str, Any]] = {}
        for line in resp_file.read_text(encoding="utf-8").splitlines():
            rec = json.loads(line)
            if rec.get("method") in methods:
                latest[rec["method"]] = rec

        for m, rec in latest.items():
            score = 0.0
            ok = False
            partial: RubPartialScore | None = None
            err: str | None = None
            try:
                pred_obj = normalize_json_text(rec["text"])
                score_res = score_exact(
                    truth, pred_obj, unordered_paths=t.unordered_paths
                )
                score = score_res.score
                ok = score_res.ok
                partial = score_partial(
                    truth, pred_obj, unordered_paths=t.unordered_paths
                )
            except Exception as exc:
                err = str(exc)

            rows.append(
                RubResultRow(
                    task_id=t.id,
                    source_case_id=t.source_case_id,
                    type=t.type,
                    track=t.track,
                    method=m,
                    model=rec.get("model"),
                    score=score,
                    partial_precision=partial.precision if partial else None,
                    partial_recall=partial.recall if partial else None,
                    partial_f1=partial.f1 if partial else None,
                    ok=ok,
                    input_tokens=int(rec.get("input_tokens", 0)),
                    output_tokens=int(rec.get("output_tokens", 0)),
                    cost_usd=float(rec.get("cost_usd", 0.0)),
                    error=err,
                )
            )

    out_csv = RUB_RESULTS_DIR / "rub_results.csv"
    write_results_csv([row.model_dump() for row in rows], out_csv)
    print(f"[green]Wrote {out_csv} ({len(rows)} rows)[/green]")


@app.command()
def rub_report() -> None:
    """Generate a RUB Markdown report from the results CSV."""
    csv_path = RUB_RESULTS_DIR / "rub_results.csv"
    if not csv_path.exists():
        raise typer.Exit(code=1)

    import pandas as pd

    df = pd.read_csv(csv_path)
    agg: dict[str, tuple[str, str]] = {
        "rus": ("score", "mean"),
        "avg_in": ("input_tokens", "mean"),
        "avg_cost": ("cost_usd", "mean"),
        "n": ("task_id", "count"),
    }
    if "partial_precision" in df.columns and df["partial_precision"].notna().any():
        agg["partial_precision"] = ("partial_precision", "mean")
    if "partial_recall" in df.columns and df["partial_recall"].notna().any():
        agg["partial_recall"] = ("partial_recall", "mean")
    if "partial_f1" in df.columns and df["partial_f1"].notna().any():
        agg["partial_f1"] = ("partial_f1", "mean")
    g = df.groupby("method").agg(**agg).reset_index()

    detail_dir = RUB_RESULTS_DIR / "detailed_reports"
    detail_dir.mkdir(parents=True, exist_ok=True)

    md_lines: list[str] = []
    md_lines.append("# RUB Report")
    md_lines.append("")
    md_lines.append(
        "This report summarizes Reconstruction Utility Benchmark (RUB) results."
    )
    md_lines.append(
        "Scores are computed on Stage B task accuracy using Markdown-only inputs."
    )
    md_lines.append("")
    md_lines.append("## Summary by method")
    md_lines.append("")
    md_lines.append(g.to_markdown(index=False))
    md_lines.append("")

    if "track" in df.columns:
        md_lines.append("## Summary by track")
        md_lines.append("")
        g_track = df.groupby(["track", "method"]).agg(**agg).reset_index()
        md_lines.append(g_track.to_markdown(index=False))
        md_lines.append("")

    for task_id, task_df in df.groupby("task_id"):
        task_path = detail_dir / f"report_{task_id}.md"
        lines = [
            "# RUB Report",
            "",
            f"## Details: {task_id}",
            "",
            task_df.to_markdown(index=False),
            "",
        ]
        task_path.write_text("\n".join(lines), encoding="utf-8")

    report_path = RUB_RESULTS_DIR / "report.md"
    report_path.write_text("\n".join(md_lines), encoding="utf-8")
    print(f"[green]Wrote {report_path}[/green]")


@app.command()
def report_public() -> None:
    """Generate chart images and update the public REPORT.md."""
    data = load_report_data()
    chart_paths = generate_charts(data)
    report_path = update_public_report(chart_paths)
    print(f"[green]Wrote {report_path}[/green]")


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
    ruleset = load_ruleset(DATA_DIR / "normalization_rules.json")
    md_outputs: dict[str, dict[str, dict[str, Any]]] = {}

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

        rules = ruleset.for_case(c.id)
        md_file = MARKDOWN_RESPONSES_DIR / f"{c.id}.jsonl"
        if md_file.exists():
            latest_md: dict[str, dict[str, Any]] = {}
            for line in md_file.read_text(encoding="utf-8").splitlines():
                rec = json.loads(line)
                if rec.get("method") in methods:
                    latest_md[rec["method"]] = rec
            md_outputs[c.id] = latest_md
        for m, rec in latest.items():
            ok = False
            score = 0.0
            score_ordered = 0.0
            score_norm: float | None = None
            score_norm_ordered: float | None = None
            score_raw: float | None = None
            score_raw_precision: float | None = None
            score_md: float | None = None
            score_md_precision: float | None = None
            err: str | None = None
            try:
                pred_obj = normalize_json_text(rec["text"])
                score = key_score(truth, pred_obj)
                score_ordered = key_score_ordered(truth, pred_obj)
                score_norm = key_score_normalized(truth, pred_obj, rules)
                score_norm_ordered = key_score_ordered_normalized(
                    truth, pred_obj, rules
                )
                score_raw = raw_coverage_score(truth, pred_obj)
                score_raw_precision = raw_precision_score(truth, pred_obj)
                md_truth = render_markdown(truth, title=c.id)
                md_rec = md_outputs.get(c.id, {}).get(m)
                if md_rec is not None:
                    md_text = str(md_rec.get("text", ""))
                    score_md = markdown_coverage_score(md_truth, md_text)
                    score_md_precision = markdown_precision_score(md_truth, md_text)
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
                    score_norm=score_norm,
                    score_norm_ordered=score_norm_ordered,
                    score_raw=score_raw,
                    score_raw_precision=score_raw_precision,
                    score_md=score_md,
                    score_md_precision=score_md_precision,
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
    if "score_norm" in df.columns:
        agg["acc_norm"] = ("score_norm", "mean")
    if "score_norm_ordered" in df.columns:
        agg["acc_norm_ordered"] = ("score_norm_ordered", "mean")
    if "score_raw" in df.columns:
        agg["acc_raw"] = ("score_raw", "mean")
    if "score_raw_precision" in df.columns:
        agg["raw_precision"] = ("score_raw_precision", "mean")
    if "score_md" in df.columns and df["score_md"].notna().any():
        agg["acc_md"] = ("score_md", "mean")
    if "score_md_precision" in df.columns and df["score_md_precision"].notna().any():
        agg["md_precision"] = ("score_md_precision", "mean")
    g = df.groupby("method").agg(**agg).reset_index()

    detail_dir = RESULTS_DIR / "detailed_reports"
    detail_dir.mkdir(parents=True, exist_ok=True)

    md_lines = []
    md_lines.append("# Benchmark Report")
    md_lines.append("")
    md_lines.append(
        "This report summarizes extraction accuracy for each method on the benchmark cases."
    )
    md_lines.append(
        "Scores are computed per case and aggregated by method. Exact, normalized, raw,"
    )
    md_lines.append(
        "and markdown tracks are reported to ensure fair comparison across variations."
    )
    md_lines.append("")
    md_lines.append("## Evaluation protocol (public)")
    md_lines.append("")
    md_lines.append("Fixed settings for reproducibility:")
    md_lines.append("")
    md_lines.append("- Model: gpt-4o (Responses API)")
    md_lines.append("- Temperature: 0.0")
    md_lines.append("- Prompt: fixed in bench/llm/openai_client.py")
    md_lines.append("- Input contexts: generated by bench.cli extract")
    md_lines.append("- Normalization: data/normalization_rules.json (optional track)")
    md_lines.append("- Evaluation: bench.cli eval (Exact + Normalized + Raw)")
    md_lines.append("- Markdown conversion: bench.cli markdown (optional)")
    md_lines.append("- Report: bench.cli report (summary + per-case)")
    md_lines.append("")
    md_lines.append("Recommended disclosure when publishing results:")
    md_lines.append("")
    md_lines.append("- Model name + version, temperature, and date of run")
    md_lines.append("- Full normalization_rules.json used for normalized scores")
    md_lines.append("- Cost/token estimation method")
    md_lines.append("- Any skipped cases and the reason (missing files, failures)")
    md_lines.append("")
    md_lines.append("## How to interpret results (public guide)")
    md_lines.append("")
    md_lines.append("- Exact: strict string match with no normalization.")
    md_lines.append(
        "- Normalized: applies case-specific rules in data/normalization_rules.json to"
    )
    md_lines.append(
        "  absorb formatting differences (aliases, split/composite labels)."
    )
    md_lines.append(
        "- Raw: loose coverage/precision over flattened text tokens (schema-agnostic)."
    )
    md_lines.append(
        "- Markdown: coverage/precision against canonical Markdown rendered from truth."
    )
    md_lines.append("")
    md_lines.append("Recommended interpretation:")
    md_lines.append("")
    md_lines.append(
        "- Use Exact to compare end-to-end string fidelity (best for literal extraction)."
    )
    md_lines.append(
        "- Use Normalized to compare document understanding across methods."
    )
    md_lines.append(
        "- Use Raw to compare how much ground-truth text is captured regardless of schema."
    )
    md_lines.append("- Use Markdown to evaluate JSON-to-Markdown conversion quality.")
    md_lines.append(
        "- When tracks disagree, favor Normalized for Excel-heavy layouts where labels"
    )
    md_lines.append("  are split/merged or phrased differently.")
    md_lines.append(
        "- Always cite both accuracy and cost metrics in public comparisons."
    )
    md_lines.append("")
    md_lines.append("## Evaluation tracks")
    md_lines.append("")
    md_lines.append("- Exact: strict string match without any normalization.")
    md_lines.append(
        "- Normalized: applies case-specific normalization rules (aliases, split/composite)"
    )
    md_lines.append(
        "  defined in data/normalization_rules.json to absorb format and wording variations."
    )
    md_lines.append(
        "- Raw: loose coverage/precision over flattened text tokens (schema-agnostic),"
    )
    md_lines.append(
        "  intended to reflect raw data capture without penalizing minor label variations."
    )
    md_lines.append(
        "- Markdown: coverage/precision comparing LLM Markdown to canonical truth Markdown."
    )
    md_lines.append("")
    md_lines.append("## Summary by method")
    md_lines.append("")
    md_lines.append(g.to_markdown(index=False))
    md_lines.append("")
    md_lines.append("## Markdown evaluation notes")
    md_lines.append("")
    md_lines.append(
        "Markdown scores measure how well the generated Markdown lines match a canonical"
    )
    md_lines.append(
        "Markdown rendering of the ground truth JSON. This is a *conversion quality*"
    )
    md_lines.append("signal, not a direct extraction-accuracy substitute.")
    md_lines.append("")
    md_lines.append("Key points:")
    md_lines.append("")
    md_lines.append(
        "- Coverage (acc_md): how much of truth Markdown content is recovered."
    )
    md_lines.append(
        "- Precision (md_precision): how much of predicted Markdown is correct."
    )
    md_lines.append(
        "- Layout shifts or list formatting differences can lower scores even if"
    )
    md_lines.append("  the underlying facts are correct.")
    md_lines.append(
        "- LLM-based conversion introduces variability; re-run with the same seed"
    )
    md_lines.append(
        "  and model settings to assess stability, or use deterministic rendering"
    )
    md_lines.append("  for baseline comparisons.")
    md_lines.append(
        "- Use Markdown scores when your downstream task consumes Markdown (e.g.,"
    )
    md_lines.append(
        "  RAG ingestion), and report alongside Exact/Normalized/Raw metrics."
    )
    md_lines.append("")
    md_lines.append("## Exstruct positioning notes (public)")
    md_lines.append("")
    md_lines.append(
        "Recommended primary indicators for exstruct positioning (RAG pre-processing):"
    )
    md_lines.append("")
    md_lines.append("- Normalized accuracy: acc_norm / acc_norm_ordered")
    md_lines.append("- Raw coverage/precision: acc_raw / raw_precision")
    md_lines.append("- Markdown coverage/precision: acc_md / md_precision")
    md_lines.append("")
    md_lines.append("Current deltas vs. best method (n=11, when available):")
    md_lines.append("")
    metric_labels = [
        ("acc_norm", "Normalized accuracy"),
        ("acc_norm_ordered", "Normalized ordered accuracy"),
        ("acc_raw", "Raw coverage"),
        ("raw_precision", "Raw precision"),
        ("acc_md", "Markdown coverage"),
        ("md_precision", "Markdown precision"),
    ]
    if "method" in g.columns and not g.empty:
        ex_row = g[g["method"] == "exstruct"]
        for metric, label in metric_labels:
            if metric not in g.columns:
                continue
            best_val = g[metric].max()
            best_methods = g[g[metric] == best_val]["method"].tolist()
            if ex_row.empty:
                ex_val = None
            else:
                ex_val = float(ex_row[metric].iloc[0])
            if ex_val is None:
                md_lines.append(f"- {label}: exstruct n/a; best {best_val:.6f}")
                continue
            delta = ex_val - best_val
            md_lines.append(
                f"- {label}: exstruct {ex_val:.6f} vs best {best_val:.6f}"
                f" ({', '.join(best_methods)}), delta {delta:+.6f}"
            )
    else:
        md_lines.append("- (summary unavailable)")
    md_lines.append("")
    md_lines.append("## Normalization leniency summary")
    md_lines.append("")
    ruleset = load_ruleset(DATA_DIR / "normalization_rules.json")
    if ruleset.cases:
        summary_rows: list[dict[str, str | int]] = []
        for case_id, rules in sorted(ruleset.cases.items()):
            details = []
            for rule in rules.list_object_rules:
                parts = [
                    f"strings={','.join(rule.string_fields) or '-'}",
                    f"strings_contains={','.join(rule.string_fields_contains) or '-'}",
                    f"lists_contains={','.join(rule.list_fields_contains) or '-'}",
                    f"strip_prefix={','.join(rule.strip_prefix.keys()) or '-'}",
                ]
                details.append(f"{rule.list_key}({'; '.join(parts)})")
            summary_rows.append(
                {
                    "case_id": case_id,
                    "alias_rules": len(rules.alias_rules),
                    "split_rules": len(rules.split_rules),
                    "composite_rules": len(rules.composite_rules),
                    "list_object_rules": len(rules.list_object_rules),
                    "details": " | ".join(details) if details else "-",
                }
            )
        md_lines.append(pd.DataFrame(summary_rows).to_markdown(index=False))
    else:
        md_lines.append("_No normalization rules defined._")
    md_lines.append("")
    md_lines.append("## Detailed reports")
    md_lines.append("")
    for case_id in sorted(df["case_id"].unique()):
        md_lines.append(f"- detailed_reports/report_{case_id}.md")
    md_lines.append("")
    out_md = RESULTS_DIR / "report.md"
    out_md.write_text("\n".join(md_lines), encoding="utf-8")
    print(f"[green]Wrote {out_md}[/green]")

    # Per-case detail reports
    detail_cols = [
        "method",
        "case_id",
        "type",
        "model",
        "score",
        "score_ordered",
        "score_norm",
        "score_norm_ordered",
        "score_raw",
        "score_raw_precision",
        "score_md",
        "score_md_precision",
        "input_tokens",
        "output_tokens",
        "cost_usd",
        "error",
    ]
    available_cols = [c for c in detail_cols if c in df.columns]

    for case_id in sorted(df["case_id"].unique()):
        case_df = df[df["case_id"] == case_id][available_cols]
        case_lines = [
            "# Benchmark Report",
            "",
            f"## Details: {case_id}",
            "",
            case_df.to_markdown(index=False),
            "",
        ]
        case_md = detail_dir / f"report_{case_id}.md"
        case_md.write_text("\n".join(case_lines), encoding="utf-8")
        print(f"[green]Wrote {case_md}[/green]")
        print(f"[cyan]Details ({case_id})[/cyan]")
        print(case_df.to_markdown(index=False))

    print("[magenta]Summary (from report.md)[/magenta]")
    print(g.to_markdown(index=False))


if __name__ == "__main__":
    app()
