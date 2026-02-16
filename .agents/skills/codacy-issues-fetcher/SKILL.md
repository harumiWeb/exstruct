---
name: codacy-issues-fetcher
description: Retrieve and format Codacy analysis issues by running `scripts/codacy_issues.py` in the ExStruct workspace. Use when users ask to inspect repository or pull-request Codacy findings, filter by severity, or produce structured issue output for review and fix planning.
---

# Codacy Issues Fetcher

Run `scripts/codacy_issues.py` as the primary interface to Codacy issue retrieval.
Avoid reimplementing API calls unless the script itself must be changed.

## Workflow

1. Confirm prerequisites.
- Run from repository root so `scripts/codacy_issues.py` is reachable.
- Ensure `CODACY_API_TOKEN` is set and valid.
- Prefer explicit `org` and `repo` if user provides them; otherwise rely on Git `origin` auto-detection.

2. Choose scope and severity.
- Repository scope: omit `--pr`.
- Pull request scope: pass `--pr <number>`.
- Severity filter (`--min-level`): `Error`, `High`, `Warning`, `Info`.
- Provider (`--provider`): `gh`, `gl`, `bb` (default is effectively `gh`).

3. Run one of the command patterns.

```powershell
# Repository issues (explicit target)
python scripts/codacy_issues.py <org> <repo> --provider gh --min-level Warning

# Pull request issues (explicit target)
python scripts/codacy_issues.py <org> <repo> --pr <number> --provider gh --min-level Warning

# Pull request issues (auto-detect org/repo from git origin)
python scripts/codacy_issues.py --pr <number> --min-level Warning
```

4. Parse output JSON and respond with actionable summary.
- Trust payload fields: `scope`, `organization`, `repository`, `pullRequest`, `minLevel`, `total`, `issues`.
- `issues` entries are formatted as:
  `<level> | <file_path>:<line_no> | <rule> | <category> | <message>`
- Report high-severity findings first, then summarize counts.

## Error Handling

- `HTTP 401` / `Unauthorized`: token invalid or missing permissions. Ask user to set or refresh `CODACY_API_TOKEN`.
- `CODACY_API_TOKEN is not set`: export the environment variable before retrying.
- `Invalid --provider`: use only `gh`, `gl`, or `bb`.
- Segment validation errors (`Invalid org/repo/pr`): sanitize input and rerun.

## Output Policy

- Return concise triage-ready results, not raw command logs.
- Include the exact command you used when reproducibility matters.
- If no issues match the selected `--min-level`, state that explicitly.
