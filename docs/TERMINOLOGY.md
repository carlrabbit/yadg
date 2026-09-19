# Terminology

## Semantic table
Stable-ID Markdown table object containing semantic header schema, body rows, optional caption, and supported inline cell content.

## Generated table
Word table whose table structure is created by YADG from a semantic table.

## Prepared table
Existing Word table whose structure/presentation is template-owned and whose repeated body rows may be populated by YADG.

## Prepared-table binding
Per-template physical placement binding from a semantic table to a prepared Word table through `{{table-rows:<id>}}`.

## Marker row
Temporary template row whose logical content is `{{table-rows:<id>}}`; it identifies the following row as the prototype and is removed from output.

## Prototype row
Word row immediately following the marker. It is cloned once per semantic body row and then removed.

## Cell placeholder
Logical `{{cell}}` text in each prototype cell indicating where the corresponding semantic source cell is authored.

## Column mapping
M0007 positional mapping between semantic Markdown columns and prototype-row Word cells.

## Template-owned header
Prepared-table row(s) displaying headings. M0007 does not populate them from Markdown header text.

## Placement override
Template-owned physical location suppressing natural-anchor rendering. Prepared-table bindings are table placement overrides.

## Authoring
Office-independent transformation of workspace values, Markdown semantics, and prepared DOCX templates into authored DOCX.

## Renderer / finalizer
Separate engine phase evaluating field/index/layout state after authoring.
