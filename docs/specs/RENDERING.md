# Rendering Specification

## Status

Authoritative for the common renderer/finalizer boundary.

Renderer-specific authority:

- Microsoft Word: `docs/specs/WORD-RENDERER.md`
- LibreOffice: retained M0005 behavior below.

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
yadg render [--workspace <path>] [--renderer <id>] [--renderer-path <path>]
```

Supported IDs after M0009:

```text
libreoffice
word
```

Default remains `libreoffice`.

`--renderer-path` applies only to LibreOffice and is invalid for Word.

## Common input/output

Renderers process top-level `YadgPreWords/*.docx` only and do not recurse.

At least one authored DOCX is required.

Renderers never modify `YadgPreWords`.

Both renderers produce:

```text
YadgWords/<name>.docx
```

Existing corresponding outputs may be replaced after preflight; unrelated files are preserved.

## Renderer-specific additional output

LibreOffice retains its existing:

```text
YadgPdfs/<name>.pdf
```

output from the same refreshed session.

The Microsoft Word renderer introduced by M0009 creates finalized DOCX only.

PDF is therefore renderer-specific, not a generic publication requirement.

## Common finalization goal

A successful renderer establishes current observable values for supported Word-native structures:

- sequence fields;
- semantic references;
- section-number references;
- template-owned indexes where supported;
- pagination-dependent state required for those displayed results.

Exact APIs differ by renderer.

## LibreOffice retained behavior

M0005 LibreOffice behavior remains:

- isolated temporary user profile;
- local UNO/API connection;
- no attachment to unrelated running instance;
- field/index refresh;
- finalized DOCX save;
- PDF export;
- bounded startup/operation;
- concrete runtime/OS provenance;
- real-runtime validation for LibreOffice claims.

## Microsoft Word specialization

Governed by `docs/specs/WORD-RENDERER.md`.

The authoritative Word locus is an interactive Windows user session with desktop Word installed.

## Failure semantics

Each renderer preflights its required runtime before corresponding normal outputs are modified.

Runtime/open/refresh/save/export failures produce actionable diagnostics and non-zero status.

Rollback of completed sibling documents is not required.

Temporary/incomplete outputs are not successful outputs.

## Publication boundary

Rendering creates finalized workspace artifacts.

Publishing is a separate M0009 command governed by `docs/specs/PUBLISHING.md`.

No renderer implicitly publishes.

## Human-visible correctness

Automated integration validates structural/observable behavior.

Milestone-scoped human artifact review applies only when explicitly required by the owning milestone.

## Deferred behavior

The common contract does not define renderer equivalence, generic PDF requirements, remote renderers, publication packages, or automatic renderer installation.
