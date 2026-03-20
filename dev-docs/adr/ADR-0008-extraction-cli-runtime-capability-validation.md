# ADR-0008: Extraction CLI Runtime Capability Validation

## Status

`accepted`

## Background

The extraction CLI currently probes Excel COM availability while building its
argument parser so that `--auto-page-breaks-dir` is shown only on COM-capable
hosts. On Windows this probe may instantiate `xlwings.App()` and launch Excel
even for lightweight commands such as `exstruct --help`.

This behavior creates two policy problems that are likely to recur:

- parser construction performs host-dependent side effects and adds startup
  latency before the user has requested any COM-only behavior
- help output changes by host capability even though the CLI syntax itself is
  stable across hosts

ExStruct already treats extraction-mode validation as part of the product
contract in ADR-0001, and treats rich-backend fallback as a runtime concern in
ADR-0002. We need a CLI-facing policy for how capability-gated extraction flags
should be exposed and validated without reintroducing startup probes.

## Decision

- Extraction CLI parser construction must remain side-effect free and must not
  probe COM or launch Excel.
- Capability-gated extraction flags may remain visible in help output when
  their syntax is stable across hosts and their requirements can be validated at
  execution time.
- `--auto-page-breaks-dir` is always exposed in extraction CLI help.
- `--auto-page-breaks-dir` is validated only when the user requests it:
  - `mode="libreoffice"` keeps the existing invalid-combination validation
  - `mode="light"` is rejected explicitly because auto page-break export
    requires COM-backed `standard` or `verbose`
  - `mode="standard"` / `mode="verbose"` require Excel COM and fail with an
    explicit runtime error when COM is unavailable
- Ordinary extraction without requested COM-only side outputs keeps the existing
  fallback policy from ADR-0002.

## Consequences

- `exstruct --help` and parser construction become faster and stop triggering
  Excel startup side effects.
- CLI help becomes consistent across hosts because it documents syntax instead
  of reflecting a startup-time environment probe.
- Users on unsupported hosts may still see `--auto-page-breaks-dir`, but they
  now receive an actionable runtime error instead of hidden syntax or silent
  export skipping.
- Future extraction CLI features that are host-capability-gated should prefer
  execution-time validation over parser-time probing unless the syntax itself is
  host-specific.

## Rationale

- Tests:
  - `tests/cli/test_cli.py`
- Code:
  - `src/exstruct/cli/main.py`
  - `src/exstruct/cli/availability.py`
- Related specs:
  - `docs/cli.md`
  - `README.md`
  - `README.ja.md`
  - `dev-docs/specs/excel-extraction.md`

## Supersedes

- None

## Superseded by

- None
