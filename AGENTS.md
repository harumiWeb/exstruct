# ExStruct AI Agents Guide

## 0. Overview

This repository is organized around the following top-level directories:

```text
exstruct/
|- src/           # Main library and implementation code
|- tests/         # Automated tests
|- sample/        # Sample workbooks and example inputs
|- schemas/       # JSON schemas and validation-related assets
|- scripts/       # Utility and maintenance scripts
|- benchmark/     # Benchmark code and performance measurements
|- docs/          # User-facing documentation
|- dev-docs/      # All developer-facing documentation
|- tasks/         # Temporary task notes and working files
|- drafts/        # Draft documents and work-in-progress materials
|- dist/          # Build artifacts and packaged outputs
`- site/          # Generated documentation site output
```

For internal development guidance, architecture notes, ADRs, specifications, and testing references, use `dev-docs/` as the canonical location. Developer-facing documentation should be written there rather than scattered across the repository.

## 1. Workflow Design

### 1. Use Plan mode by default

- Always start tasks with 3 or more steps, or tasks that affect architecture, in Plan mode
- If things stop going well partway through, do not force it; stop immediately and replan
- Use Plan mode not only for implementation, but also for verification steps
- Write detailed specifications before implementation to reduce ambiguity

### 2. Multi-Agent Strategy

- Actively use sub-agents to keep the main context window clean
- Delegate research, investigation, and parallel analysis to sub-agents
- For complex problems, use sub-agents to apply more compute resources
- To keep execution focused, assign one task per sub-agent
- Use explorer for read-heavy codebase exploration
- Use worker for implementation and fixes
- Use reviewer for reviews

### 3. Self-Improvement Loop

- Whenever you receive a correction from the user, record that pattern in `tasks/lessons.md`
- Write rules for yourself so you do not repeat the same mistake
- Keep improving those rules thoroughly until the error rate goes down
- At the start of each session, review the lessons relevant to the project

### 4. Always verify before completion

- Do not mark a task as complete until you can prove that it works
- Compare the main branch and your changes when necessary
- Ask yourself, "Would a staff engineer approve this?"
- Run tests, review logs, and show that it works correctly

### 5. Pursue elegance (with balance)

- Before making an important change, pause and ask, "Is there a more elegant way to do this?"
- If a fix feels hacky, think, "Based on everything I know now, implement an elegant solution"
- Skip this process for simple and obvious fixes (do not over-engineer)
- Question your own work before presenting it

### 6. Autonomous bug fixing

- When you receive a bug report, fix it directly without needing step-by-step guidance
- Use logs, errors, and failing tests to solve it yourself
- Eliminate context switching for the user
- Even without being asked, go fix failing CI tests

---

## 2. Areas Outside the AI's Responsibility (Handled by Humans)

The AI does not own the following areas. Humans make these decisions.

- Specification decisions (the direction of ExStruct's evolution)
- Public API design (deciding whether something is a breaking change)
- Large-scale reorganization of the directory structure
- Security and licensing decisions

However, the AI **may make proposals**.

---

## 3. Required Work Procedure

The AI must always follow the steps below before generating code.

1. **Understand requirements**: Read specifications and design materials, and fully understand the requirements
2. **Consider the design**: Consider function decomposition and model design as needed.
3. **Define the specification**: Based on the requirements, define function argument and return types in `tasks/feature_spec.md`.
4. **Assign tasks**: Clearly define each task and determine the implementation order.
5. **Implement code**: Implement the code while following the standards above.
6. **Review code**: Self-review generated code and confirm that it meets the quality standards.
7. **Generate tests**: Generate test code as needed.
8. **Run tests**: Run the generated test code and confirm that it behaves as expected.
9. **Static analysis**: Run `uv run task precommit-run` and confirm that there are no mypy / Ruff errors.
10. **Update documentation**: If there are changes, update the related documentation as well.

---

## 4. Task Management

1. **Plan first**: Write the plan in `tasks/todo.md` as checkable items
2. **Review the plan**: Review it before starting implementation
3. **Track progress**: Mark completed items as you go
4. **Explain changes**: Provide a high-level summary at each step
5. **Document results**: Add a Review section to `tasks/todo.md`
6. **Record lessons**: Update `tasks/lessons.md` after receiving corrections

---

## 5. Documentation Retention Policy

### Separation of Roles

- `tasks/todo.md` may temporarily hold not only session-specific progress tracking, but also verification results, unresolved items, and summaries of decision rationale.
- `tasks/feature_spec.md` may be used as a pre-implementation working spec draft, but do not treat it as disposable if it contains specifications, constraints, or validation conditions that will be referenced in the future.
- `tasks/lessons.md` is where recurrence-prevention rules are stored, and should not be used to store design decisions or the specification itself.
- Permanent internal documentation belongs under `dev-docs/`.
- Move design decisions and trade-offs to `dev-docs/adr/`, current internal specifications and constraints to `dev-docs/specs/`, and implementation structure and extension guidance to `dev-docs/architecture/`.
- Only user-facing contracts such as public API, CLI, and MCP should be reflected in the corresponding documents under `docs/`.

### Using skills

- If you are unsure where to store a document, where to move it, or how to verify it, prefer using available skills over relying on manual judgment alone.
- Use `adr-suggester` to determine whether an ADR is needed, `adr-drafter` for ADR drafts or update proposals, `adr-linter` to lint drafts, `adr-reviewer` for design review, `adr-reconciler` for drift audits, and `adr-indexer` for index synchronization.
- Do not leave skill results trapped in temporary notes under `tasks/`; reflect them in the appropriate `dev-docs/` or `docs/` location as needed.

### Information to Keep

- Decision rationale that future implementers may encounter again on the same issue
- Chosen policies adopted after comparing multiple options
- Permanent rules established through review, CI, Codacy, or incident response
- Contracts related to public API, CLI, MCP, output formats, validation, and compatibility
- Specification context behind added regression tests where forgetting the reason could cause the issue to recur

### Information You May Discard

- One-off notes about work order
- Rejected hypotheses or interim notes that ended midway
- Progress logs with no reference value after completion
- Simple lists of steps with no decision rationale

### Required Steps at Completion

- At task completion, review the relevant sections of `tasks/feature_spec.md` and `tasks/todo.md`, and classify each item as either "temporary notes that can be discarded", "content that should remain in a permanent spec", or "content that should remain as an ADR".
- The AI must not blank out all of `tasks/feature_spec.md` or `tasks/todo.md` based on its own judgment. Cleanup must be limited to the relevant sections of the completed task.
- If there is content that will be referenced in the future, move it into permanent documentation before deleting anything.
- Do not discard decision rationale, specifications, or validation conditions before migration is complete.
- Only sections confirmed to contain no permanent information may be summarized, deleted, or archived.
- If ADR creation, spec creation, index synchronization, or design review is involved, and a corresponding skill exists, run it first and use its verdict and findings to decide the permanent document destination and what to reflect there.
- Choose the destination according to the role split defined in `dev-docs/README.md`.
- Prefer `dev-docs/adr/` for "why", `dev-docs/specs/` for "what is guaranteed", and `dev-docs/architecture/` for "how the structure works".
- Only when the change affects a public contract should you update the corresponding page under `docs/` in addition to moving the information into internal documentation.

### When to Create an ADR

- If you are unsure whether an ADR is needed, first use `adr-suggester` to determine `required` / `recommended` / `not-needed` and record the rationale.
- If any of the following apply, the AI must record the decision under `dev-docs/adr/`:
  - There are trade-offs or a comparison between multiple options.
  - The same question may recur in the future.
  - The design intent cannot be understood from the implementation diff alone.
  - A permanent policy was established through review, CI, Codacy, or incident investigation.
  - It is highly likely to be referenced by later implementation or review.

### End-of-Session Checklist

- Confirm that conclusions in the Review section of `tasks/todo.md` have been moved, as needed, into `dev-docs/adr/`, `dev-docs/specs/`, `dev-docs/architecture/`, or `docs/`.
- Confirm that contracts, constraints, and validation conditions in `tasks/feature_spec.md` have been reflected, as needed, in permanent documents under `dev-docs/`.
- If an ADR was added / updated / superseded, confirm as needed that the results of `adr-linter`, `adr-reviewer`, `adr-reconciler`, and `adr-indexer` do not conflict with the permanent documents.
- Only after the information has been moved into permanent documentation may the relevant sections be shortened.

---

## 6. Core Principles

- **Simplicity first**: Keep every change as simple as possible. Minimize the code affected.
- **No cutting corners**: Find the root cause. Avoid temporary fixes. Maintain senior engineer standards.
- **Minimize impact**: Limit changes to only what is necessary. Do not introduce new bugs.
