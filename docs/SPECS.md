# YADG Product Specification

## Purpose

YADG makes mandatory Word-based documentation maintainable as version-controlled Markdown without losing the structure and presentation supplied by prepared Microsoft Word templates.

The product is template-first rather than a general-purpose DOCX layout generator.

## Product authority model

1. Markdown is authoritative for maintainable document content and semantic document objects.
2. `YADG.md` is authoritative for workspace-level scalar values and workspace notes.
3. Prepared DOCX templates are authoritative for structure, presentation, presentation vocabulary, numbering conventions, and Word-native structures unless explicitly delegated.
4. Generated DOCX/PDF files are derived artifacts.
5. Office-independent authoring is authoritative for semantic resolution and OOXML transformation.
6. Layout-dependent field evaluation/finalization belongs to a renderer.

## Processing model

YADG:

1. parses workspace metadata and Markdown into an OOXML-free semantic model plus workspace values;
2. authors DOCX through tags, value substitution, template-local bindings, semantic references, and OOXML structures;
3. renders/finalizes layout-dependent state with a concrete renderer.

Stages 1 and 2 require neither Microsoft Office nor LibreOffice.

## Workspace and CLI

A workspace is rooted at `YADG.md`.

`YADG.md` may contain optional machine-readable front matter followed by human-facing notes. It is not ordinary Markdown document source.

```text
yadg check [--workspace <path>]
yadg build [--workspace <path>]
yadg render [--workspace <path>] [--renderer libreoffice] [--renderer-path <path>]
```

Templates are top-level `YadgTemplates/*.docx`; authored outputs are `YadgPreWords/*.docx`; M0005 finalized outputs are `YadgWords/*.docx` and `YadgPdfs/*.pdf`.

## Identity

Semantic object IDs are case-sensitive and workspace-wide:

```text
[A-Za-z][A-Za-z0-9_-]*
```

Workspace value IDs use the same lexical form but occupy a separate namespace.

Display text, labels, captions, filenames, source positions, bookmark names, and field codes are not semantic identity.

## Template vocabulary

```text
{{content:<stable-id>}}
{{section:<stable-id>}}
{{table:<stable-id>}}
{{figure:<stable-id>}}
{{value:<value-id>}}
```

`content`, `section`, `table`, and `figure` are block operations.

`value` is the inline scalar substitution operation defined in `docs/specs/WORKSPACE-VALUES.md`.

Template-control/prototype markers are defined in `docs/specs/WORD-REFERENCES.md`.

## Markdown references

```markdown
See Figure [@system-context].
See Table [@interface-matrix].
See Section [@architecture].
```

`[@id]` resolves semantic objects only and emits the numeric target result.

Workspace values are not interpolated into Markdown in M0006.

## Structured content and Word references

Structured content is governed by `docs/specs/STRUCTURED-CONTENT.md`.

Caption/reference/template-prototype behavior is governed by `docs/specs/WORD-REFERENCES.md`.

## Workspace values

Workspace scalar values come only from optional YAML front matter in root `YADG.md`.

M0006 values are case-sensitive, single-line strings in a namespace separate from semantic object IDs.

Template `{{value:<id>}}` tags substitute them literally while preserving template-owned surrounding formatting.

The schema, supported Word locations, formatting rule, and validation behavior are authoritative in `docs/specs/WORKSPACE-VALUES.md`.

## Template-local metadata

Template front matter maps semantic roles to existing template styles/prototypes. It is presentation binding, not workspace content/value authority.

## Check and build

`check` validates the complete workspace and each template without outputs.

M0006 additionally validates workspace front matter and all supported template value references.

`build` performs equivalent validation before modifying normal outputs.

Value replacement is complete during Office-independent authoring and requires no renderer evaluation.

## Renderer boundary

M0005 implements LibreOffice as the first renderer under `docs/specs/RENDERING.md`.

The renderer evaluates layout/field/index state and produces finalized DOCX/PDF artifacts.

Workspace values require no renderer-specific semantics.

## Diagnostics

Diagnostics are first-class and should identify stable code, source/template location, offending object/tag/value/binding, and corrective context where feasible.

## Non-goals

YADG is not a general-purpose Word layout engine or generic Markdown-to-DOCX converter, does not require Office for ordinary authoring, and does not treat generated artifacts as source authority.
