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

M0010 establishes realistic document composition:

- Markdown headings are rebased relative to the template outline context at each `section`/`content` insertion anchor;
- effective Word heading levels 1 through 9 are template-style roles;
- ordinary Markdown paragraphs retain template-owned paragraph formatting from the placement anchor;
- workspace values substitute visible template text across supported Word stories, including notes, comments, and text boxes.

Detailed behavior is in `docs/specs/DOCUMENT-COMPOSITION.md`.

## Template vocabulary

Block placement/control vocabulary remains main-document-body oriented.

M0010 does not add block-placement tags to footnotes, endnotes, comments, headers, footers, or text boxes.

Template front-matter heading bindings support effective levels 1 through 9.

## Rendering

Supported renderers:

```text
libreoffice
word
```

LibreOffice remains the default for compatibility.

Both renderers produce finalized `YadgWords/*.docx`.

LibreOffice retains its existing PDF side output.

Microsoft Word remains the V1 fidelity target.

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

Publishing never invokes build/render and does not publish PDFs.

## Compatibility proof

Before V1 release-readiness, M0010 requires committed realistic DOCX fixtures originating independently from Microsoft Word and LibreOffice Writer.

Both fixtures must pass YADG authoring and both supported renderer paths.

These application-produced fixtures complement, rather than replace, focused synthetic OpenXML tests.

## Platform

Office-independent authoring components remain isolated from Microsoft Word automation.

The complete Word-enabled product's authoritative build/integration locus is Windows.

The M0010 full compatibility target additionally requires real Microsoft Word and LibreOffice.

## Diagnostics

Diagnostics remain first-class.

Invalid heading rebasing, unsupported effective heading depth, malformed/missing values in supported stories, and realistic-template incompatibilities fail before successful authoring is claimed.

## Non-goals

YADG is not a DMS client, generic Word layout engine, general plugin host, or generic Markdown-to-DOCX converter.

M0010 does not introduce Markdown syntax for Word footnotes, endnotes, comments, text boxes, headers, or footers.
