# YADG Product Specification

## Purpose

YADG is a template-first document authoring tool: Markdown and `YADG.md` own maintainable content/data while prepared DOCX templates own structure and presentation.

## Authority model

1. Markdown owns semantic document objects and table data.
2. `YADG.md` owns workspace scalar values/notes.
3. DOCX templates own presentation, Word-native structures, and prepared-table layout unless explicitly delegated.
4. Generated DOCX/PDF are derived artifacts.
5. Office-independent authoring owns semantic resolution and OOXML transformation.
6. Rendering owns layout-dependent evaluation.

## CLI/artifacts

```text
yadg check [--workspace <path>]
yadg build [--workspace <path>]
yadg render [--workspace <path>] [--renderer libreoffice] [--renderer-path <path>]
```

Templates: `YadgTemplates/*.docx`; authored: `YadgPreWords/*.docx`; finalized: `YadgWords/*.docx`; PDFs: `YadgPdfs/*.pdf`.

## Template vocabulary

Existing tags remain:

```text
{{content:<id>}}
{{section:<id>}}
{{table:<id>}}
{{figure:<id>}}
{{value:<id>}}
```

M0007 adds prepared-table control syntax:

```text
{{table-rows:<table-id>}}
{{cell}}
```

These controls are valid only under `docs/specs/PREPARED-TABLES.md`.

## Tables

One Markdown pipe table remains one semantic table. It may be rendered as a generated Word table or populate a prepared template-owned Word table.

Prepared mode preserves template table/header/layout/style structures and clones only the designated prototype row. Markdown header defines semantic column order/count; Word header rows provide visible presentation.

Caption/reference semantics remain attached to the same semantic table.

## Values, references, rendering

Workspace values follow `docs/specs/WORKSPACE-VALUES.md`; semantic references/captions follow `docs/specs/WORD-REFERENCES.md`; renderer finalization follows `docs/specs/RENDERING.md`.

Prepared row population is complete during `build`; the renderer does not populate rows.

## Validation

`check` validates all workspace/template semantics without outputs. `build` performs equivalent validation before modifying normal outputs.

M0007 additionally validates prepared-table marker/prototype structure, positional column mapping, forbidden prototype content, placement uniqueness, and caption/reference interaction.

## Non-goals

YADG is neither a general Word layout engine nor a generic Markdown-to-DOCX converter and does not treat generated artifacts as source authority.
