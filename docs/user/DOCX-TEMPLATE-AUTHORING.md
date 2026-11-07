# DOCX template authoring

This is the canonical guide for creating and maintaining prepared Word templates for YADG.

Prepared DOCX templates in `YadgTemplates/` own presentation, layout, numbering, fields, and other Word-native structure. Markdown owns maintainable semantic content.

For source syntax, see [Markdown authoring](MARKDOWN-AUTHORING.md).

## Core principle

YADG intentionally reuses the template rather than recreating Word formatting from scratch.

Prefer a real template resource or concrete example whenever Word stores important presentation in structures that cannot be reconstructed reliably from a style name alone.

The template should own:

- page and section layout;
- headers and footers;
- typography;
- corporate styles;
- numbering appearance;
- TOC/list-of-figures/list-of-tables fields;
- caption sequence conventions;
- prepared table layout;
- presentation examples/prototypes.

Markdown should own semantic document content.

## Visible text tags

YADG uses ordinary visible text tags rather than Word content controls.

Examples:

<!-- M0015-example: template-visible-tags -->
```text
{{section:architecture}}
{{content:architecture}}
{{table:interfaces}}
{{figure:system-context}}
{{value:document-version}}
```

This keeps the template understandable in Word and lets authors see the intended insertion points and formatting context directly.

## Section and content placement

```text
{{section:id}}
```

inserts the selected Markdown heading plus its body.

```text
{{content:id}}
```

inserts only the body below that heading.

YADG rebases source heading hierarchy relative to the template insertion context. The template owns the effective Word heading presentation and numbering.

## Table and figure placement

Identified semantic tables and figures can be placed directly:

```text
{{table:interfaces}}
{{figure:system-context}}
```

Direct placement suppresses the corresponding natural Markdown emission for that template.

Built-in text diagrams and Mermaid diagrams both reach template authoring as figure-like content when they have stable IDs. The template can place either through the same:

```text
{{figure:id}}
```

mechanism. The DOCX template does not define text-diagram geometry and does not run Mermaid.

Anonymous auto-detected text diagrams have no stable identity and therefore render only at their natural source position.

## Workspace values in visible template text

Use:

```text
{{value:document-version}}
```

for source-controlled values defined in `YADG.md`.

YADG replaces supported visible text while preserving formatting from the relevant template run/text position.

Value substitution can apply across the supported Word stories, including main body, tables, headers, footers, notes, comments, and supported text-box/shape text.

It does not rewrite arbitrary package metadata, field instruction code, relationship targets, or other non-visible XML.

## Template front matter

Template-local YADG configuration is visible YAML between marker paragraphs:

```text
{{yadg:frontmatter}}
```

and:

```text
{{/yadg:frontmatter}}
```

Example:

<!-- M0015-example: template-frontmatter-schema -->
```yaml
version: 1

styles:
  headings:
    1: "Heading 1"
    2: "Heading 2"
  lists:
    unordered: "ListBullet"
    ordered: "ListNumber"
  generatedTable: "Normal Table"
  caption: "Normal"
  codeInline: "Code Character"
  blockquote: "Quote"
  codeBlock: "HTML Preformatted"

prototypes:
  figureCaption: figure-caption
  tableCaption: table-caption
  unorderedListItem: bullet-item
  orderedListItem: number-item
```

Use only keys supported by the current template schema.

For Markdown blockquotes and fenced or indented code blocks, the selected template must contain concrete paragraph styles for `blockquote` and `codeBlock`. The defaults are `Quote` and `HTMLPreformatted`; bind these roles to other template paragraph styles when needed. Code blocks use the selected preformatted style for every source line and do not inherit the placement marker's paragraph style. Nested quotes keep the quote style and gain indentation at each nested level.

## Presentation resolution principle

The author-facing model is:

> when YADG needs to create Word content, prefer a concrete template example/resource over reconstructing presentation from hidden metadata.

Presentation may come from:

- the paragraph containing a placement tag;
- a named prototype;
- a Word paragraph/character/table style;
- Word numbering;
- a prepared table row;
- the run containing a value/cell placeholder;
- a compatibility fallback.

Styles are one mechanism inside that model, not the whole model.

## Inspecting a template

### Inspect styles

```powershell
yadg inspect styles
yadg inspect styles --template report.docx
```

This lists concrete serialized Word styles and relevant metadata such as:

- style type;
- Word display name;
- internal style ID;
- aliases;
- visibility flags;
- numbering capability where applicable.

Word's UI name and internal style ID are not always the same.

Strict selectors resolve in this order:

1. exact internal style ID;
2. case-insensitive exact primary display name;
3. case-insensitive exact alias.

Matching is not fuzzy. Ambiguous names/aliases are errors.

### Inspect effective template resolution

```powershell
yadg inspect template
yadg inspect template --template report.docx
```

This answers:

> What will YADG actually use from this template?

It reports placements, prototypes, semantic presentation roles, resolved sources, human-searchable Word locations, and YOLO fallback candidates.

## Headings

Heading presentation is resolved through configured/template heading roles.

YADG understands effective outline semantics, including inherited Word style relationships, rather than assuming that every valid heading must literally be named `Heading1`, `Heading2`, and so on.

When source content is inserted below an existing template outline context, the source hierarchy is rebased relative to that location.

## List styles

Without a prototype, a list may use a concrete paragraph style whose effective numbering resolves to the required semantic kind.

Unordered lists require bullet numbering.

Ordered lists require real ordered Word numbering. Decimal, Roman, alphabetic, and other valid ordered formats can qualify.

A Word gallery label does not prove that a usable serialized style/numbering definition exists. If strict validation rejects a style, inspect the actual template resources.

## List-item prototypes

A prototype is usually the most reliable way to preserve a corporate list format.

Create a real bulleted/numbered paragraph in Word and place it between:

```text
{{yadg:prototype:bullet-item}}
```

and:

```text
{{/yadg:prototype:bullet-item}}
```

The prototype contains exactly one logical:

```text
{{item}}
```

Bind it:

<!-- M0015-example: template-list-prototype-binding -->
```yaml
prototypes:
  unorderedListItem: bullet-item
```

The visible markers and placeholder form the prototype body:

<!-- M0015-example: template-list-prototype -->
````text
{{yadg:prototype:bullet-item}}
{{item}}
{{/yadg:prototype:bullet-item}}
````

The prototype owns its Word numbering, indentation, paragraph properties, and run formatting. YADG clones that concrete example for each Markdown item.

For an ordered-list prototype, create a real numbered Word paragraph and use the same placeholder pattern with its configured ID:

<!-- M0015-example: template-ordered-list-prototype -->
````text
{{yadg:prototype:number-item}}
{{item}}
{{/yadg:prototype:number-item}}
````

## Caption prototypes and references

Numeric captions require real Word field structure.

A configured figure/table caption prototype contains:

- exactly one `SEQ` field;
- exactly one logical `{{caption}}`;
- no other YADG authoring/control placeholder.

The prototype owns label text, sequence identifier, punctuation, paragraph/run formatting, and style.

YADG clones the prototype, replaces `{{caption}}`, and establishes the corresponding bookmark/reference mechanics.

Without truthful numeric caption structure, YADG does not invent a number.

Create visible marker paragraphs around one paragraph containing the real Word field and caption placeholder. In Word, insert a `SEQ Figure` field at the bracketed position:

<!-- M0015-example: template-caption-prototype -->
````text
{{yadg:prototype:figure-caption}}
Figure [SEQ Figure field]: {{caption}}
{{/yadg:prototype:figure-caption}}
````

The table equivalent uses `{{yadg:prototype:table-caption}}`, a real `SEQ Table` field, and `{{caption}}`.

<!-- M0015-example: template-table-caption-prototype -->
````text
{{yadg:prototype:table-caption}}
Table [SEQ Table field]: {{caption}}
{{/yadg:prototype:table-caption}}
````

This applies equally to ordinary image figures, identified built-in text diagrams, and Mermaid figures because all are figure semantics at the template boundary.

Root `YADG.md` may select how generated `[@id]` references display:

```yaml
references:
  headings: number
  figures: label-number
  tables: label-number
```

Headings support `number`; figures and tables support `number` or `label-number`. Omitted settings preserve numeric-only legacy output. In `label-number` mode, YADG derives the label and spacing/punctuation before the `SEQ` field from the selected caption prototype. Use the label text and Word field format already present in that real template prototype; YADG does not supply English/German labels. A prototype with no clear label before its single sequence field cannot support this mode.

Ordinary Markdown links such as `[the model description](#liquidity-model)` use a Word bookmark for the explicitly identified heading, figure, or table that actually renders. Their authored wording stays visible and clickable. A template using `{{content:liquidity-model}}` does not emit the heading itself and therefore cannot be the target of a heading fragment link.

## Prepared tables

When the Word template owns the desired table layout, use a prepared row marker:

```text
{{table-rows:interfaces}}
```

The immediately following Word row is the prototype row.

Each prototype cell contains exactly one ordinary paragraph and one logical:

```text
{{cell}}
```

YADG clones the prototype row once per Markdown body row and preserves template-owned widths, borders, shading, and formatting.

Prepared-table placement is an alternative to generated-table placement for the same semantic table.

## Inline code presentation

If the workspace uses:

```yaml
markdown:
  codeInline: style
```

configure a character style:

```yaml
styles:
  codeInline: "Code"
```

`codeInline` requires a character style, not a paragraph style. Since inline code uses `style` by default, a template used with inline code must provide this role (or the workspace may explicitly choose `ignore` or `error`). Nested lists use the template's existing ordered/unordered paragraph roles and Word numbering levels up to level nine. Blockquotes retain surrounding template paragraph presentation; YADG does not add quotation labels or decorative marks.

## Generated tables

Generated tables use the configured/default concrete table style where available.

Under YOLO, YADG may select a compatible table resource or preserve semantic table content without the requested style rather than dropping it.

## TOC, indexes, fields, and renderer finalization

The template owns Word-native structures such as TOCs, lists of figures/tables, fields, page numbering, headers/footers, and similar layout-dependent constructs.

YADG authors semantic content into the prepared DOCX. Word or LibreOffice finalization then updates renderer-dependent state according to the selected renderer's contract.

Do not try to encode these Word-native presentation structures in Markdown merely because the source content is version-controlled.

## YOLO presentation fallbacks

Strict mode uses the explicitly defined template-resolution chain and errors when required presentation cannot be resolved.

YOLO first attempts strict resolution and may then use:

1. a compatible related template resource;
2. a versioned built-in fallback;
3. plain truthful presentation.

Every fallback is reported as a degradation.

YOLO does not make an ambiguous template resource authoritative and does not invent semantic numbering/reference identity.

## Human-searchable Word diagnostics

DOCX files do not have useful source line numbers.

Where possible YADG reports context such as:

```text
YadgTemplates/report.docx
body > "Operational Risks" > paragraph 27
near: "Availability of the upstream system"
```

Use the `near:` text with Word Find, and use `yadg inspect template` to understand the surrounding YADG resource/placement context.

## Template author checklist

Before treating a template as production-ready:

- keep visible YADG tags intentional and unique where required;
- run `yadg inspect styles` when selecting styles by name/ID;
- run `yadg inspect template` to verify effective resources;
- use prototypes for presentation that depends on real numbering/field/layout structure;
- verify caption prototypes before relying on numeric figure/table references;
- ensure prepared-table marker/prototype rows are valid;
- retain Word-native TOC/index/field structures in the template;
- run strict `yadg check` and `yadg build` with representative source;
- finalize with the renderer required by the document workflow.

## Related guides

- [Getting started](GETTING-STARTED.md)
- [Markdown authoring](MARKDOWN-AUTHORING.md)
- [Troubleshooting](TROUBLESHOOTING.md)
