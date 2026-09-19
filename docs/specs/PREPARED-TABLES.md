# Prepared Table Population Specification

## Status

Authoritative for prepared Word table row population introduced by M0007.

## Purpose

Generated Markdown tables remain available when a template delegates complete table construction to YADG. Prepared-table population serves the complementary case: the Word template already owns the table structure, column widths, borders, row/cell formatting, visible header labels, and surrounding layout while Markdown owns semantic table data.

## Semantic source

The source is an existing stable-ID Markdown pipe table. Its header defines semantic column count/order; its body rows are repeated data. For prepared-table population the Markdown header is not rendered into Word: template-owned rows provide visible headings.

## Template binding

A prepared Word table binds to a semantic table through a marker row whose only non-whitespace logical text across all cells is:

```text
{{table-rows:<table-id>}}
```

The immediately following row in the same table is the prototype row.

The marker row and prototype row are authoring controls and are removed from output.

## Placement semantics

A prepared-table binding is a physical placement override for that semantic table.

Per template, a semantic table may have at most one physical override: either one `{{table:<id>}}` generated-table direct placement or one `{{table-rows:<id>}}` prepared-table binding. Using both, or multiple prepared bindings, is a `check` error.

A prepared binding suppresses natural-anchor generated-table emission for that template.

## Prototype row and zero rows

YADG clones the prototype once per Markdown body row, in source order, inserts the clones at the prototype position, then removes marker/prototype controls.

If the Markdown table has zero body rows, marker and prototype are removed and no data row is emitted.

Template-owned rows before/after remain unchanged.

## Column mapping

Mapping is positional.

The prototype row must contain exactly the same number of physical cells as the semantic table has columns.

Each prototype cell must contain exactly one ordinary paragraph and exactly one logical:

```text
{{cell}}
```

The placeholder may be split across runs. Literal template text may appear before/after it in the same paragraph.

## Cell authoring

The corresponding semantic source cell replaces `{{cell}}`.

Supported source-cell inlines are the existing table-cell subset: text, emphasis, strong, soft break, hard break, and semantic numeric reference.

The run containing the first logical character of `{{cell}}` provides base run properties. Markdown emphasis/strong augments that base formatting. Literal prefix/suffix text remains unchanged.

## Prototype restrictions

The prototype row must not contain another YADG tag, pre-existing bookmarks/fields, drawings/images, nested tables, content controls, horizontal grid-span merges, or vertical merges.

Ordinary row/cell/paragraph/run presentation properties, widths, shading, borders, and related non-identity formatting are cloned.

## Header ownership

Prepared-table population does not rewrite template header rows. Markdown header text is semantic schema only for this mode. M0007 does not compare visible Word header labels with Markdown headers.

## Caption and references

The semantic table retains existing optional caption/reference semantics.

When prepared placement is the unique physical placement of a captioned table, YADG emits the existing caption structure immediately after the prepared Word table using existing caption-prototype rules. That caption remains the numbered/reference target.

M0007 does not bind to a separately pre-existing template caption paragraph.

## Workspace values

M0006 value substitution continues in ordinary template-owned rows. The M0007 prototype row itself may contain only `{{cell}}` YADG placeholders; workspace-value tags inside the prototype row are deferred.

## Validation

`check` validates marker syntax/type, physical placement uniqueness, marker location/content, immediate prototype existence, column count, one paragraph/placeholder per prototype cell, forbidden prototype structures, and existing caption/reference eligibility.

`build` performs equivalent validation before normal output modification.

## Preservation

Except for intentional control-row replacement, YADG preserves table properties/style/grid, column widths, template-owned rows, cloned prototype formatting, and unrelated document structure. It does not recalculate prepared-table widths or synthesize presentation.

## Compatibility

Existing natural/generated table behavior and `{{table:<id>}}` remain unchanged when no prepared binding exists.

## Deferred behavior

M0007 does not define named/header-based mapping, automatic header population, merged prototype cells, multiple row variants, nested/subrows, prototype images, value tags inside prototype rows, pre-existing caption binding, duplicate table placement, external row sources, filtering/sorting/grouping, or totals/subtotals.
