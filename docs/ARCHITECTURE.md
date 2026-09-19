# Architecture

## Objective

YADG separates workspace/content semantics, Word package transformation, template-owned presentation/prototypes, and rendering.

## `authoring-core`

Owns workspace discovery/values, Markdown parsing, OOXML-free semantic sections/tables/figures, stable IDs/references, and Office-independent validation.

A semantic table contains header schema, body rows, optional caption, and stable identity independent of Word.

## `word-authoring`

Owns template tags/front matter/prototypes, value substitution, generated tables/figures, prepared-table analysis, per-template placement planning, row cloning/population, styles, captions, bookmarks, and fields.

OOXML row/cell types do not leak into core semantic models.

## Prepared-table adapter

```text
semantic table -> generated Word table
```

or:

```text
semantic table
+ existing Word table
+ marker row
+ prototype row
-> populated existing Word table
```

Prepared binding participates in the same per-template physical placement plan as direct `{{table:id}}`.

## Renderer

The renderer owns field/index/layout finalization only. Prepared-row population is not deferred to LibreOffice or Word.

## Preservation

Prepared-table authoring clones template row/cell/paragraph/run presentation structures and replaces only control placeholders with semantic source-cell content. Table widths, borders, header presentation, style, and unrelated rows remain template authority.

## Runtime

M0007 introduces no new external runtime, filesystem, credential, or network boundary.
