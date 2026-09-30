# Templates and presentation

Prepared DOCX templates in `YadgTemplates/` own presentation and Word-native structure.

YADG intentionally reuses the template rather than trying to recreate Word formatting from scratch.

## Visible template tags

YADG uses ordinary visible text tags.

Examples:

```text
{{section:architecture}}
{{content:architecture}}
{{table:interfaces}}
{{figure:system-context}}
{{value:document-version}}
```

This keeps templates understandable in Word and avoids requiring authors to manage content controls.

## Template front matter

Template-specific YADG configuration is stored as visible YAML between marker paragraphs:

```text
{{yadg:frontmatter}}
```

and:

```text
{{/yadg:frontmatter}}
```

Example:

```yaml
version: 1

styles:
  headings:
    1: "Heading 1"
    2: "Heading 2"
  lists:
    unordered: "Corporate Bullet"
    ordered: "Corporate Number"
  generatedTable: "Table Grid"
  caption: "Caption"
  codeInline: "Code"

prototypes:
  figureCaption: figure-caption
  tableCaption: table-caption
  unorderedListItem: bullet-item
  orderedListItem: number-item
```

Use only keys supported by the current template schema.

## Presentation principle

The author-facing model is:

> when YADG needs to create Word content, prefer a concrete template example/resource over reconstructing presentation from hidden metadata.

Presentation may come from:

- the paragraph containing a placement tag;
- a named prototype;
- a Word paragraph/character/table style;
- Word numbering;
- a prepared table row;
- the run containing a value/cell placeholder;
- a compatibility default.

Styles are one mechanism inside that model, not the whole model.

## Inspect the template

### Styles

```powershell
yadg inspect styles
yadg inspect styles --template report.docx
```

This lists concrete serialized Word styles and reports information such as:

- style type;
- Word display name;
- internal style ID;
- aliases;
- visibility flags;
- numbering capability where applicable.

Word's UI name and internal ID are not always the same. Renaming a style can change the visible name while leaving the ID unchanged.

Strict selectors resolve in this order:

1. exact internal style ID;
2. case-insensitive exact primary display name;
3. case-insensitive exact alias.

Matching is not fuzzy.

An ambiguous name/alias is an error.

### Effective template resolution

```powershell
yadg inspect template
yadg inspect template --template report.docx
```

This answers a different question:

> What will YADG actually use from this template?

It reports placements, prototypes, semantic presentation roles, resolved sources, and YOLO fallback candidates.

A valid list prototype, for example, is reported as the actual strict presentation source instead of separately complaining that an unused list style is missing.

## Headings

Heading presentation is resolved through the configured/template heading roles.

YADG understands effective outline semantics, including inherited Word style relationships, rather than assuming that every valid heading must literally be named `Heading1`, `Heading2`, etc.

When inserting a selected Markdown section/content block, YADG rebases source heading hierarchy relative to the prepared template context.

The template owns the resulting Word heading presentation/numbering.

## List styles

Without a prototype, list presentation can use a concrete paragraph style whose effective numbering resolves to the correct semantic kind.

Unordered lists require bullet numbering.

Ordered lists require real ordered Word numbering. They are not restricted to decimal: Roman/alphabetic and other valid ordered formats are allowed.

A Word gallery name by itself does not prove that a concrete serialized style/numbering definition exists.

If strict validation says a list style has no usable numbering, run:

```powershell
yadg inspect styles
```

and choose a real concrete resource.

## List prototypes

A prototype is often the easiest way to preserve a corporate/template-specific list format.

Create a real bulleted/numbered paragraph in Word and place it between marker paragraphs:

```text
{{yadg:prototype:bullet-item}}
```

and:

```text
{{/yadg:prototype:bullet-item}}
```

The prototype block contains exactly one Word paragraph with exactly one logical:

```text
{{item}}
```

Bind it:

```yaml
prototypes:
  unorderedListItem: bullet-item
```

For an ordered list:

```yaml
prototypes:
  orderedListItem: number-item
```

The prototype owns real Word numbering, indentation, paragraph properties and run formatting. YADG clones it for each Markdown item.

Prototype presentation takes precedence over an otherwise unnecessary list-style binding.

## Caption prototypes

Numeric captions require real Word field structure.

A configured figure/table caption prototype contains:

- exactly one `SEQ` field;
- exactly one logical `{{caption}}`;
- no other YADG authoring/control placeholder.

The prototype owns label text, sequence identifier, punctuation, paragraph/run formatting and style.

YADG clones the prototype, replaces `{{caption}}`, and creates internal bookmark/reference mechanics.

Without truthful numeric caption structure, YADG does not invent a number.

## Prepared tables

When the Word template already owns the table layout, bind a Markdown table to a prepared table with a marker row:

```text
{{table-rows:interfaces}}
```

The immediately following Word row is the prototype row.

Each prototype cell contains exactly one ordinary paragraph and exactly one logical:

```text
{{cell}}
```

YADG clones the prototype row once per Markdown body row and preserves template-owned widths, borders, shading and formatting.

Prepared-table placement is an alternative to generated-table placement for the same semantic table.

## Inline code style

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

`codeInline` requires a character style, not a paragraph style.

Under YOLO, if the configured style is unavailable, YADG can use a compatible code/source/monospace character style or preserve the text normally.

## Generated tables

Generated tables use the configured/default concrete table style where available.

Under YOLO, YADG may choose another compatible table style or preserve the semantic table without a style reference rather than dropping its content.

## Values and formatting

For visible value replacement, the existing template run containing the logical value tag provides the base formatting example.

YADG replaces content while preserving the surrounding template-owned formatting contract.

## YOLO presentation fallbacks

Strict mode uses the explicitly defined template resolution chain and errors when required presentation cannot be resolved.

YOLO first tries strict resolution, then may use:

1. a compatible related template resource;
2. a versioned built-in fallback;
3. plain truthful presentation.

Built-in list identities include:

```text
builtin:unordered-list-v1
builtin:ordered-list-v1
```

These use real Word numbering rather than literal bullet/number text.

Fallback selection is deterministic for identical template/input/YADG version.

Every fallback is reported as a degradation.

## Human-searchable Word diagnostics

Word documents do not have useful source line numbers.

Where possible, YADG reports:

```text
YadgTemplates/report.docx
body > "Operational Risks" > paragraph 27
near: "Availability of the upstream system"
```

Use the `near:` text with Word Find.

The heading path uses effective heading semantics, including inherited styles.

## What the template should own

Prefer putting these things in the template:

- page/section layout;
- headers/footers;
- typography;
- corporate styles;
- numbering appearance;
- TOC/list-of-figures/list-of-tables fields;
- caption sequence conventions;
- prepared table layout;
- concrete presentation examples/prototypes.

Prefer putting maintainable semantic content in Markdown/YADG configuration.

That boundary is the central YADG design.
