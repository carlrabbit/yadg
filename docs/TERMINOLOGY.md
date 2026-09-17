# Terminology

This document defines YADG terms used as project authority.

## Authoring

The Office-independent transformation of Markdown and prepared Word templates into structurally complete DOCX artifacts through parsing, semantic resolution, validation, and OOXML manipulation.

## Renderer / finalizer

A separate executable or runtime component that opens an authored DOCX in a real rendering engine and updates layout-dependent state such as fields, references, indexes, table of contents data, pagination-related values, and optional PDF output.

The initial concrete renderer target is Microsoft Word through Office Interop on Windows.

## Template

A prepared DOCX document that remains authoritative for document structure and presentation unless a specific template location explicitly delegates a structure element to Markdown.

## Template-owned structure

Headings, cover pages, section ordering, styles, captions, tables, and other document structure intentionally present in the Word template.

## Markdown-owned content

Maintainable source content stored in Markdown and version control. Markdown content is selected by stable semantic references.

## Tag

A visible textual YADG marker embedded in a DOCX template. Tags are deliberately ordinary document text rather than Word content controls.

Tag recognition must operate on logical paragraph/cell text and cannot assume that one tag maps to one OOXML text run.

## Stable ID

An explicit semantic identifier used to reference a Markdown object. Display text, heading text, caption text, position, and filenames are not stable semantic identity unless a specification explicitly says otherwise.

## Section

A Markdown heading together with the content structurally belonging to that heading up to the next heading of the same or higher level.

## Section content

The body belonging to a section, excluding the referenced section heading itself but including nested subsections.

## Derived artifact

A generated output such as an authored DOCX, finalized DOCX, PDF, validation report, or evidence bundle. Derived artifacts are rebuildable outputs, not project source authority.

## Validation target

The concrete system whose behavior is being established by a validation path, for example the OOXML package representation or a locally installed Microsoft Word runtime.

## Validation locus

Where validation executes, for example ordinary local/CI .NET execution or a Windows machine with Microsoft Word installed.
