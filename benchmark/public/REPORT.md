# Benchmark Summary (Public)

This summary consolidates the latest results for the Excel document benchmark and
RUB (structure query track). Use this file as a public-facing overview and link
full reports for reproducibility.

Sources:
- outputs/results/report.md (core benchmark)
- outputs/rub/results/report.md (RUB structure_query)
<!-- CHARTS_START -->
## Charts

![Core Benchmark Summary](outputs/plots/core_benchmark.png)
![Markdown Evaluation Summary](outputs/plots/markdown_quality.png)
![RUB Structure Query Summary](outputs/plots/rub_structure_query.png)
<!-- CHARTS_END -->
## Scope

- Cases: 12 Excel documents
- Methods: exstruct, openpyxl, pdf, html, image_vlm
- Model: gpt-4o (Responses API)
- Temperature: 0.0
- Note: record the run date/time when publishing
- This is an initial benchmark (n=12) and will be expanded in future releases.

## Core Benchmark (extraction + scoring)

Key metrics from outputs/results/report.md:

- Exact accuracy (acc): best = pdf 0.607551, exstruct = 0.583802
- Normalized accuracy (acc_norm): best = pdf 0.856642, exstruct = 0.835538
- Raw coverage (acc_raw): best = exstruct 0.876495 (tie for top)
- Raw precision: best = exstruct 0.933691
- Markdown coverage (acc_md): best = pdf 0.700094, exstruct = 0.697269
- Markdown precision: best = exstruct 0.796101

Interpretation:
- pdf leads in Exact/Normalized, especially when literal string match matters.
- exstruct is strongest on Raw coverage/precision and Markdown precision,
  indicating robust capture and downstream-friendly structure.

## RUB (structure_query track)

RUB evaluates Stage B questions using Markdown-only inputs. Current track is
"structure_query" (paths selection).

Summary from outputs/rub/results/report.md:

- RUS: exstruct 0.166667 (tie for top with openpyxl 0.166667)
- Partial F1: exstruct 0.436772 (best among methods)

Interpretation:
- exstruct is competitive for structure queries, but the margin is not large.
- This track is sensitive to question design; it rewards selection accuracy
  more than raw reconstruction.

## Positioning for RAG/LLM Preprocessing

Practical strengths shown by the current benchmark:
- High Raw coverage/precision (exstruct best)
- High Markdown precision (exstruct best)
- Near-top normalized accuracy

Practical caveats:
- Exact/normalized top spot is often pdf
- RUB structure_query shows only a modest advantage

Recommended public framing:
- exstruct is a strong option when the goal is structured reuse (JSON/Markdown)
  for downstream LLM/RAG pipelines.
- pdf/VLM methods can be stronger for literal string fidelity or visual layout
  recovery.

## Known Limitations

- Absolute RUS values are low in some settings (task design sensitive).
- Results vary by task type (forms/flows/diagrams vs tables).
- Model changes (e.g., gpt-4.1) require separate runs and reporting.

## Next Steps (optional)

- Add a reconstruction track that scores “structure rebuild” directly.
- Add task-specific structure queries (not only path selection).
- Publish run date, model version, and normalization rules with results.
