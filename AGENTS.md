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

## 3. Task Management

Before generating or modifying code, perform the following steps according to the scale of your work:

1. Understand the requirements: Review relevant specification documents, ADR documentation, and existing implementations.
2. Consider the design implications: Assess impact scope, compatibility with current designs, and alternative approaches.
3. If necessary, create working notes:
   - For recurrence prevention: `tasks/lessons.md`
4. Add or update tests as needed.
5. Implement changes.
6. Verify functionality.
7. Run tests.
8. Conduct self-review.
9. Update documentation, ADR documents, specifications, and the CHANGELOG as appropriate.

- Any updates to ADR documents or specifications must be recorded in the respective directories:
- For ADR documents: `dev-docs/adr/`
- For specification documents: `dev-docs/specs/`
- If changes affect public APIs, they may require recording in the following documentation:
- Specification documents within `dev-docs/specs/`
- Overview descriptions in the `README.md` file

---

## 4. Documentation Retention Policy

### Separation of Roles

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

- If there is content that will be referenced in the future, move it into permanent documentation before deleting anything.
- Do not discard decision rationale, specifications, or validation conditions before migration is complete.
- Only sections confirmed to contain no permanent information may be summarized, deleted, or archived.
- If ADR creation, spec creation, index synchronization, or design review is involved, and a corresponding skill exists, run it first and use its verdict and findings to decide the permanent document destination and what to reflect there.
- Choose the destination according to the role split defined in `dev-docs/README.md`.
- Prefer `dev-docs/adr/` for "why", `dev-docs/specs/` for "what is guaranteed", and `dev-docs/architecture/` for "how the structure w
