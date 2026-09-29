# Rendering Specification

## Status

Authoritative for the common renderer/finalizer boundary after the M0012 correction.

Renderer-specific authority:

- Microsoft Word: `docs/specs/WORD-RENDERER.md`
- LibreOffice: retained renderer behavior below.

## Boundary

```text
YadgTemplates + Markdown
        |
        | yadg build
        v
YadgPreWords/*.docx
        |
        | yadg render
        v
renderer-specific finalized artifacts
```

`render` never parses Markdown, resolves authoring tags, or implicitly invokes `build`.

## CLI

```text
yadg render [--workspace <path>] [--renderer <id>] [--renderer-path <path>] [--yolo]
```

Supported IDs:

```text
libreoffice
word
```

Default remains `libreoffice`.

`--renderer-path` applies only to LibreOffice and is invalid for Word.

## Common input

Renderers process top-level:

```text
YadgPreWords/*.docx
```

only and do not recurse.

At least one authored DOCX is required.

Renderers never modify PreWords.

## Generated output ownership

The finalized workspace directories are YADG-owned generated artifact sets.

### Actual LibreOffice renderer

A successful invocation produces the complete current set:

```text
YadgWords/<current-preword-name>.docx
YadgPdfs/<current-preword-name>.pdf
```

for every current top-level PreWord.

No stale finalized DOCX/PDF from an earlier render remains after successful completion.

### Actual Microsoft Word renderer

A successful invocation produces the complete current set:

```text
YadgWords/<current-preword-name>.docx
```

for every current top-level PreWord.

Because Word does not produce PDF, a successful Word render leaves no stale YADG-generated PDF from a previous LibreOffice render in `YadgPdfs`.

### Renderer fallback

When YOLO falls back to the alternate renderer, output ownership follows the **actual** renderer, not the requested renderer.

## Result-set commit

The renderer may stage individual document processing however it chooses, but successful workspace commit is a complete result-set replacement.

The required invariant is:

```text
success
=> generated finalized directories reflect this invocation only
```

Do not preserve unrelated/stale files inside YADG-generated artifact directories.

Do not delete the prior committed result set merely because a new renderer attempt has begun.

A failed renderer attempt does not claim a new finalized set.

Preserving the previous committed result set after failure is preferred and is required for the staged YOLO fallback path.

## Common finalization goal

A successful renderer establishes current observable values for supported Word-native structures:

- sequence fields;
- semantic references;
- section-number references;
- template-owned indexes where supported;
- pagination-dependent state required for displayed results.

Exact APIs differ by renderer.

## LibreOffice retained behavior

LibreOffice retains:

- isolated temporary user profile;
- local UNO/API connection;
- no attachment to unrelated running instance;
- field/index refresh;
- finalized DOCX save;
- PDF export;
- bounded startup/operation;
- concrete runtime/OS provenance;
- real-runtime validation.

## Microsoft Word specialization

Governed by `docs/specs/WORD-RENDERER.md`.

The authoritative Word locus is an interactive Windows user session with desktop Word installed.

## Failure semantics

Each renderer preflights its required runtime before committing a new finalized result set.

Runtime/open/refresh/save/export failures produce actionable diagnostics and non-zero status.

Temporary/incomplete outputs are not successful outputs.

## Explicit renderer fallback under YOLO

Strict `render` uses only the requested renderer.

`render --yolo` may try the other supported renderer after the requested renderer fails.

Each attempt operates on staged copies of the authored input; failed attempt outputs are discarded before fallback.

If both fail, render fails and authored intermediates remain intermediates.

Successful fallback reports explicitly:

```text
requested renderer
actual renderer
actual runtime version
YOLO fallback = yes
```

The actual renderer identity is carried explicitly in renderer/result state and is not inferred from nullable executable/path fields.

One renderer substitution is one logical degradation event; primary diagnostics may be attached as details.

## Publication boundary

Rendering creates finalized workspace artifacts.

Publishing is separate and governed by `PUBLISHING.md`.

No renderer implicitly publishes.

## Validation additions

Prove at minimum:

1. render A+B, reduce PreWords to A, render again -> Words contains only A;
2. LibreOffice A+B then A -> both Words and PDFs contain only A;
3. LibreOffice success followed by Word success -> stale PDFs are absent;
4. YOLO fallback commits exactly the actual renderer's result set;
5. failed primary+fallback does not delete the previously committed finalized result set merely because a new attempt started.

## Human-visible correctness

Automated integration validates structural/observable behavior.

Milestone-scoped human artifact review applies only when explicitly required by the owning milestone.

## Deferred behavior

The common contract does not define renderer equivalence, generic PDF requirements, remote renderers, publication packages, or automatic renderer installation.
