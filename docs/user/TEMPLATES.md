# Templates

Prepared DOCX templates in `YadgTemplates/` control layout and Word-native structures. YADG uses visible text tags such as `{{content:introduction}}`; it does not create styles or replace tags with content controls.

## Discover styles

Run:

```powershell
yadg inspect styles
yadg inspect styles --template report.docx
```

This reads concrete styles from the DOCX directly, even if ordinary `check` fails for another reason. Output groups styles by type and reports the Word display name, internal ID, aliases, visibility flags when present, and paragraph numbering capability.

Word distinguishes a style's internal ID from its primary UI name and aliases. Renaming a style in Word can change its visible name without changing its internal ID. Selectors match an exact ID first, then a case-insensitive exact primary name, then an exact alias. Matching is not fuzzy. An ambiguous name/alias or a style of the wrong type is an error. Hidden and semi-hidden styles may exist without appearing as ordinary gallery choices. A gallery option may be latent rather than a concrete style in the document.

Template front matter remains version 1. For example:

```yaml
version: 1
styles:
  headings:
    1: "Heading 1"
  lists:
    unordered: "List Bullet 2"
    ordered: "List Number"
  generatedTable: "Table Grid"
  caption: "Caption"
  codeInline: "Code"
prototypes:
  unorderedListItem: bullet-item
  orderedListItem: number-item
```

Use `yadg inspect styles` to find names and aliases. Existing style IDs remain valid. Headings and list/caption roles require paragraph styles, generated tables require a table style, and `codeInline` requires a character style.

## Lists

Without a list-item prototype, configure paragraph styles whose Word numbering relationship is usable. If `check` reports `List style '<name>' does not resolve to a valid numbering definition`, inspect the style list and choose the actual name/alias of a concrete list style. Word's visible gallery can include latent styles, and a style name alone does not guarantee numbering is associated.

A template prototype is the explicit escape hatch. Create a real Word/Writer numbered or bulleted paragraph, place it between `{{yadg:prototype:<id>}}` and `{{/yadg:prototype:<id>}}`, and use exactly one `{{item}}` in that one paragraph. Bind it under `prototypes.unorderedListItem` or `prototypes.orderedListItem`. The prototype paragraph owns numbering and formatting. YADG clones it per item; it does not synthesize bullet or number glyphs.

## Other template controls

Front matter begins with a paragraph containing `{{yadg:frontmatter}}`, declares `version: 1`, and ends at `{{/yadg:frontmatter}}`. Existing caption prototypes, prepared table row markers, bookmarks, and fields remain template-owned. See the focused specifications linked from [the spec index](../SPECS.md) for exact template contracts.
