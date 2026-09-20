# YADG Product Specification

## Purpose

YADG maintains document content as source-controlled Markdown/workspace data while prepared Word DOCX templates remain authoritative for document presentation and Word-native structure.

YADG authors, finalizes, and explicitly publishes finalized Word documents.

## Authority model

1. Markdown owns maintainable semantic document content.
2. `YADG.md` owns workspace values, producer configuration, workspace notes, and optional publication defaults.
3. DOCX templates own presentation and Word-native structures unless explicitly delegated.
4. External content producers supply constrained generated semantic products.
5. Office-independent authoring produces structurally complete authored DOCX.
6. Renderers finalize fields/indexes/layout-dependent document state.
7. Publishing copies finalized DOCX from workspace state to an explicit delivery destination.

## CLI

After M0009:

```text
yadg check [--workspace <path>]
yadg build [--workspace <path>]
yadg render [--workspace <path>] [--renderer libreoffice|word] [--renderer-path <path>]
yadg publish [--workspace <path>] [--publish-path <path>]
```

`--renderer-path` is LibreOffice-specific.

## Workspace artifacts

```text
YadgTemplates/*.docx   source templates
YadgPreWords/*.docx    authored intermediate documents
YadgWords/*.docx       finalized documents
YadgPdfs/*.pdf         LibreOffice-specific existing PDF output
```

Published DOCX files live at the effective publication destination rather than in a mandatory workspace publication directory.

## Authoring

Existing Markdown sections/content, figures, tables, prepared tables, references, values, and external Mermaid producers remain in force.

`check`/`build` remain semantically independent of rendering.

## Rendering

Supported renderers:

```text
libreoffice
word
```

LibreOffice remains the default for compatibility.

Both renderers produce finalized `YadgWords/*.docx`.

LibreOffice retains its existing PDF side output.

The Microsoft Word renderer is the V1 fidelity target and is specified in `docs/specs/WORD-RENDERER.md`.

## Publishing

`publish` copies all top-level finalized `YadgWords/*.docx` to an explicit filesystem destination.

Effective destination:

```text
--publish-path
    overrides
YADG.md publish.path
    otherwise error
```

Relative paths resolve against workspace root.

Publishing never invokes build/render and does not publish PDFs in M0009.

Detailed behavior is in `docs/specs/PUBLISHING.md`.

## Platform

Office-independent authoring components remain isolated from Microsoft Word automation.

The complete Word-enabled product's authoritative build/integration locus is Windows.

M0009 does not promise preservation of a Linux full-solution build.

## Diagnostics

Diagnostics remain first-class.

Word renderer diagnostics include actionable automation-stage/COM context.

Publish diagnostics identify source/destination failures without modifying document content.

## Non-goals

YADG is not a DMS client, generic Word layout engine, general plugin host, or generic Markdown-to-DOCX converter.
