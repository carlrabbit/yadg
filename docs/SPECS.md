# YADG Product Specifications

## Purpose

This file is the index and cross-cutting map for active product specifications.

Detailed normative behavior lives in focused files under `docs/specs/`.

When a focused specification explicitly states that it supersedes an earlier rule, the newer focused specification governs that overlapping subject.

## Core authoring and composition

- `docs/specs/AUTHORING-EXPERIENCE.md` — workspace bootstrap, configurable Markdown authoring policies, diagnostics, `check --list`, and authoring-oriented CLI behavior.
- `docs/specs/DOCUMENT-COMPOSITION.md` — template-relative headings, paragraph composition, visible value substitution, and realistic document compatibility.
- `docs/specs/STRUCTURED-CONTENT.md` — lists, tables, figures, assets, captions, and float-like placement; list-style details are specialized/superseded by `TEMPLATE-STYLES.md`.
- `docs/specs/PREPARED-TABLES.md` — prepared Word table-row population.

## Template presentation and Word structures

- `docs/specs/TEMPLATE-STYLES.md` — style discovery/resolution, style role types, Word UI name/alias handling, inline-code character styles, numbering resolution, and list-item prototypes.
- `docs/specs/WORD-REFERENCES.md` — template control region/front matter, caption prototypes, bookmarks/fields, and semantic numeric references. Its older raw-style-ID requirement is superseded by `TEMPLATE-STYLES.md`.

## Workspace/configuration and producers

- `docs/specs/WORKSPACE-VALUES.md` — `YADG.md`, workspace values, and schema-v1 root configuration.
- `docs/specs/CONTENT-PRODUCERS.md` — constrained external content producers and Mermaid.

## Rendering and delivery

- `docs/specs/RENDERING.md` — renderer-neutral finalization contract.
- `docs/specs/WORD-RENDERER.md` — Microsoft Word renderer specialization.
- `docs/specs/PUBLISHING.md` — finalized DOCX publication behavior.

## Cross-cutting product model

YADG remains template-first:

```text
Markdown/workspace data own maintainable semantic content.
Prepared DOCX templates own presentation and Word-native structure.
YADG authors Office-independent DOCX.
A selected renderer finalizes fields/index/layout-dependent state.
Publishing copies finalized DOCX explicitly.
```

YADG 1.1 improves authoring/discovery ergonomics without changing that ownership model.

## Presentation and authoring resilience authority

- `docs/specs/TEMPLATE-PRESENTATION.md` — generalized template-owned examples/resources and deterministic strict/YOLO presentation role resolution; supersedes narrower style/list presentation framing where stated.
- `docs/specs/AUTHORING-RESILIENCE.md` — explicit `--yolo`, recovery catalog, visible degradation output, renderer fallback, and fatal boundaries; supersedes earlier failure-only statements for the named recoverable cases.
