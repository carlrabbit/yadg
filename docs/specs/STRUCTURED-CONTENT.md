# Structured Content Specification

## Status

Authoritative for M0003 and later structured-content behavior unless superseded by a later project decision.

## Purpose

This specification defines YADG semantics for lists, generated Markdown tables, figures/images, semantic captions, asset validation, and float-like table/figure placement.

The central rule is:

> Tables and figures retain their semantic source position in Markdown. A direct table/figure tag in a particular Word template explicitly relocates that semantic object for that generated document.

Placement is therefore explicit authoring behavior, not implicit deduplication.

## Semantic structured objects

M0003 adds these semantic block concepts:

- unordered list;
- ordered list;
- table;
- figure;
- natural structured-object anchor.

Tables and figures are first-class referenceable objects.

M0016 adds ordinary `[authored words](#stable-id)` links to rendered tables and figures, alongside generated numeric `[@stable-id]` references. An ordinary fragment link preserves its words and targets the one actual emitted occurrence; it does not request a number or caption text. See `WORD-REFERENCES.md` for workspace display settings and target validation.

A table or figure carries one stable ID that participates in the workspace-wide registry shared with sections and future stable objects.

The semantic model must not contain Open XML SDK types.

## Lists

### Supported source form

M0017 supports unordered and ordered Markdown lists, including nested ordered/unordered combinations up to nine effective Word numbering levels. Static task-list markers are also supported.

Examples:

```markdown
- First
- Second with **strong** text
```

```markdown
1. First
2. Second
```

Each list item contains one paragraph using the supported inline subset and may contain nested child lists. Nested lists preserve source order and use the template's numbering at the corresponding level.

YADG does not support:

- multi-paragraph list items;
- block content other than nested lists inside list items.

Unsupported list forms fail `check`.

### Word rendering

List presentation, selectors, and the list-item prototype escape hatch are specialized by `docs/specs/TEMPLATE-STYLES.md`, which supersedes the fixed-style-only list binding below. `ListBullet` and `ListNumber` remain the default compatibility selectors when no front-matter binding is supplied.

Unordered list items use the configured existing Word paragraph style (default internal ID):

```text
ListBullet
```

Ordered list items use the configured existing Word paragraph style (default internal ID):

```text
ListNumber
```

The template owns the visual appearance and numbering configuration.

For a list style required by selected content, `check` must establish that:

1. the required style exists;
2. its effective paragraph numbering resolves to a valid numbering definition in the template.

YADG does not generate substitute bullet/number glyphs when the template numbering contract is absent.

M0003 does not synthesize arbitrary list appearance.

Nested items use the same template-owned list role/prototype with the effective Word numbering level set to the nesting depth. Depths beyond nine are diagnosed before output mutation. Task markers are static `☐` / `☒` glyphs in the list paragraph, not Word content controls.

## Tables

### Markdown syntax and identity

M0017 uses Markdig's `Table`, `TableRow`, and `TableCell` AST as the sole semantic source for pipe and Pandoc-style grid tables. YADG does not split Markdown table lines or cell content itself.

Pipe tables are the preferred form for simple rectangular tables. Grid tables are the alternative when cells need multiple lines or merged rows/columns. Each table requires one `{#id}` attribute-only line immediately after the table; optional `caption="..."` remains supported on that line. The metadata line is attached to the table and never rendered as a paragraph.

A table becomes a referenceable YADG table only when immediately followed by an attribute-only line containing its stable ID:

```markdown
| System | Protocol |
|---|---|
| Alpha | HTTPS |
| Beta | SFTP |
{#interface-matrix}
```

The attribute line belongs to the table and is not rendered.

Every table must have a stable ID. An un-ID'd Markdown table is invalid.

The attribute line accepts a stable ID and the existing optional caption attribute.

### Cell content

Table cells support:

- plain text;
- emphasis;
- strong emphasis;
- soft and hard line breaks;
- multiple paragraphs in grid-table cells;
- Markdig-represented horizontal and vertical grid-cell spans.

Cells do not support lists, headings, nested tables, fenced code, images, or other non-paragraph blocks. Unsupported cell blocks fail with an actionable diagnostic. Prepared-table row population continues to accept rectangular single-paragraph cells; merged and multi-paragraph shapes are not valid prepared-row input.

Pipe-table column alignment markers may be parsed but M0003 does not establish a Word alignment contract from them; Word presentation is template-owned.

### Word rendering

Generated tables use the existing Word table style ID:

```text
TableGrid
```

`check` fails when a selected table requires this style and the style is absent.

The Markdown header row is semantically marked as the first/header row in OOXML, and the generated table enables first-row table-style behavior.

The table uses automatic Word table layout rather than YADG calculating final column widths. Grid cell paragraphs remain separate Word paragraphs. Markdig `ColumnSpan` and `RowSpan` map to OOXML `gridSpan` and `vMerge`; invalid or inconsistent coverage is rejected rather than guessed.

### M0017 additional blocks and inline forms

Markdig AST blocks now support fenced and indented literal code (except existing diagram-fence handling), nested ordered/unordered lists through nine effective numbering levels, and blockquotes containing paragraphs or nested blockquotes. List items support one paragraph plus nested lists. Task markers in list items are static checkbox glyphs. GFM strikethrough maps to Word strike formatting.

The M0003 exclusions below for nested lists and the M0012 exclusions below for ordinary code fences are superseded by M0017. Multi-paragraph list items, unrelated block types in list items, and unsupported block types in quotes or grid cells remain errors. Inline code styling defaults to the template-owned `codeInline` character role; explicit `ignore` and `error` policies remain available.

M0003 does not populate or mutate a prepared existing Word table. `table-rows` or equivalent prepared-table behavior is explicitly postponed.

### Table captions

Table captions are not defined in M0003.

Caption numbering, `SEQ` fields, bookmarks, cross-references, and list-of-tables behavior belong to the later references/document-structures milestone.

## Figures

### Markdown syntax and identity

A M0003 figure is a block paragraph consisting of one Markdown image with one stable ID attached directly to the image:

```markdown
![System context](images/system-context.png){#system-context}
```

The image alt text is the figure's semantic caption text.

An empty alt text means the figure has no caption:

```markdown
![](images/logo.png){#logo}
```

Every M0003 figure must have a stable ID.

Ordinary inline images and images without stable IDs are unsupported.

Markdown image titles are not part of the M0003 figure contract and must not be used as caption or metadata.

### Supported image formats

M0003 supports:

- PNG;
- JPEG/JPG.

Other image formats fail `check`.

### Asset path resolution

The image target is a filesystem path relative to the Markdown source file containing the figure.

Remote URLs and data URIs are unsupported.

The resolved asset must:

- exist as a readable file;
- have a lexical path within the declared workspace;
- have a supported extension/decodable format.

The selected local workspace is trusted input. A file link is usable when normal OS file access can read its content; YADG does not resolve its target to enforce physical containment. Lexical `../` escapes remain invalid. Recursive Markdown discovery does not follow linked child directories.

Path normalization and validation happen before output artifacts are modified.

### Image size

M0003 uses deterministic authored image dimensions:

1. derive intrinsic image width/height from pixel dimensions;
2. map pixels at 96 DPI to physical OOXML extent;
3. never upscale above intrinsic size;
4. if intrinsic width exceeds the effective text width at the final placement anchor, scale width down to that text width and scale height proportionally.

The effective text width is determined from the Word section/page dimensions and margins governing the final placement location.

If a valid effective text width cannot be determined for a figure that needs downscaling, `check` fails rather than guessing a layout width.

The later renderer may affect pagination but must not be required to make the image fit the authored text width contract.

### Figure caption

When alt text is non-empty, YADG emits a caption paragraph immediately after the image paragraph using the existing Word paragraph style ID:

```text
Caption
```

`check` fails if a rendered non-empty figure caption requires `Caption` and the template lacks that style.

M0003 emits caption text only. It does not insert `Figure N`, `SEQ`, bookmark, cross-reference, or list-of-figures field structures.

When alt text is empty, no caption paragraph is generated.

## Natural anchors and float-like placement

### Natural source position

Every table and figure has a natural anchor at the place where the object appears in Markdown.

When selected through `{{content:id}}` or `{{section:id}}`, the object normally renders at that source anchor.

The object remains semantically part of its surrounding section regardless of where a Word template ultimately places it.

### Direct placement tags

A Word template may explicitly take responsibility for physical placement:

```text
{{table:interface-matrix}}
{{figure:system-context}}
```

A direct structured-object tag means:

> Render this semantic object at this template location and suppress its natural source-anchor rendering for this template.

This is a placement override, not an additional copy.

Placement overrides are resolved per template before section/content rendering.

### Placement invariants

For each table or figure and each template:

- **zero direct placement tags** — render at every natural anchor reached through rendered section/content selections;
- **exactly one direct placement tag** — render at the direct template location and suppress all natural-anchor occurrences for that object in that template;
- **more than one direct placement tag** — `check` error.

A direct placement tag may render an object even if no selected `content`/`section` region would otherwise reach its natural anchor.

A direct tag must resolve to an object of the matching type. For example, `{{table:x}}` cannot resolve a figure or section.

There is no implicit cross-type or cross-tag deduplication.

M0003 does not provide a clone/duplicate-render operation. If a future use case requires deliberate repeated rendering of one structured object, it requires an explicit later semantic contract.

### Word tag placement

`table` and `figure` tags are block tags.

They must be the only non-whitespace logical content of a supported main-document-body paragraph and remain subject to the same split-run logical-text parsing rules as existing tags.

The direct-tag paragraph acts as the insertion anchor.

Generated table/figure blocks replace that anchor.

## Interaction with section/content selections

Structured objects remain blocks inside the semantic section body.

For example:

```markdown
## Architecture {#architecture}

Intro.

![System context](images/context.png){#system-context}

Details.
```

Without a direct figure tag:

```text
{{content:architecture}}
```

renders:

```text
Intro.
[figure]
[caption if present]
Details.
```

With:

```text
{{content:architecture}}
...
{{figure:system-context}}
```

the section content renders:

```text
Intro.
Details.
```

and the figure renders at the direct template location.

The semantic meaning of `content:architecture` is unchanged. The per-template placement plan controls where the figure block is physically emitted.

## Validation and build ordering

`check` must be able to build a complete per-template placement plan before authoring.

At minimum it validates:

- object IDs and global uniqueness;
- table/figure source syntax;
- asset paths/formats;
- required template styles/numbering;
- type-correct direct references;
- direct-placement uniqueness;
- supported tag locations;
- figure sizing prerequisites.

`build` performs equivalent validation before modifying normal output artifacts.

## Deferred structured-content behavior

M0003 deliberately does not define:

- prepared Word table row population;
- table captions;
- figure/table numbering fields;
- bookmarks;
- cross-references;
- list of figures/tables;
- floating layout chosen automatically by Word/YADG without an explicit template placement;
- arbitrary image resize attributes;
- SVG/GIF/TIFF/BMP support;
- inline images;
- nested lists;
- custom per-workspace style mappings.

Those require later project authority.

## M0012 resilience specialization

The rules above describe strict mode. Under explicit `--yolo`, missing in-workspace figure files may become visible ID/caption/path placeholders, and failed Mermaid execution may become a visible semantic figure placeholder. Lexical path escapes, unsupported asset kinds, malformed Markdown identity, and other semantic errors remain fatal. A link attribute alone is not an error. See `AUTHORING-RESILIENCE.md`.
