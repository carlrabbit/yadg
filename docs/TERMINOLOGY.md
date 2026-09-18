# Terminology

This document defines YADG terms used as project authority.

## Authoring

The Office-independent transformation of Markdown and prepared Word templates into structurally complete DOCX artifacts through parsing, semantic resolution, validation, and OOXML manipulation.

## Renderer / finalizer

A separate component that opens an authored DOCX in a real rendering engine and updates layout-dependent state such as fields, references, indexes, TOCs, pagination-related values, and optional PDF output.

The initial concrete renderer target is Microsoft Word through Office Interop on Windows.

## Workspace

One YADG documentation unit processed by one `check` or `build` invocation.

A workspace is rooted at a directory containing `YADG.md`.

## Workspace marker

The root-level `YADG.md` file identifying a YADG workspace.

Its body is optional human-facing documentation and is not document source content.

## Template

A prepared DOCX document authoritative for structure and presentation unless a specific location delegates structure or placement to YADG.

## Template-owned structure

Headings, cover pages, section ordering, styles, captions/style definitions, tables, and other structure intentionally present in the Word template.

## Markdown-owned content

Maintainable source content and semantic objects stored in Markdown and version control.

## Tag

A visible textual YADG marker embedded in a DOCX template.

Tags are ordinary document text, not Word content controls, and are parsed from logical paragraph text across OOXML run boundaries.

## Block tag

A tag whose replacement produces one or more Word block elements.

Current block tags must be the only non-whitespace logical content of their Word paragraph.

## Stable ID

An explicit, case-sensitive semantic identifier with syntax:

```text
[A-Za-z][A-Za-z0-9_-]*
```

Stable IDs are workspace-wide across all referenceable object types.

## Reference registry

The workspace-wide mapping from stable IDs to referenceable semantic objects, including sections, tables, and figures.

## Section

A Markdown heading with a stable ID together with content belonging to that heading up to the next heading of the same or higher level.

## Section content

The section body excluding the referenced section heading while retaining nested headings and blocks.

## Structured object

A semantic block with independent stable identity and render behavior.

M0003 structured objects are tables and figures.

## Natural anchor

The semantic source position of a structured object in Markdown.

Without a placement override, rendering a selection containing the anchor renders the object at that position.

## Placement override

A unique direct `table` or `figure` tag in a Word template that takes responsibility for the physical placement of that semantic object in the generated document.

A placement override suppresses natural-anchor rendering for that object in that template.

It is relocation, not duplication.

## Figure

A block image semantic object with stable identity, an image asset, and optional caption text.

In M0003, figure caption text comes from Markdown image alt text.

## Derived artifact

A generated output such as an authored DOCX, finalized DOCX, PDF, validation report, or evidence bundle.

Derived artifacts are rebuildable outputs, not source authority.

## Validation target

The concrete system whose behavior establishes required evidence, such as a real DOCX/OOXML package or installed Microsoft Word runtime.

## Validation locus

Where validation executes, such as an ordinary local/CI .NET environment or a Windows machine with Microsoft Word installed.
