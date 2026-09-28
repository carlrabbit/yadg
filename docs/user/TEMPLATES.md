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

## Template presentation and fallback preview

Run `yadg inspect template` to see front matter, placements, prototypes, role bindings, resolved style IDs/names/types, and searchable nearby Word text. The report previews deterministic YOLO candidates when strict role resolution fails. `yadg inspect styles` remains the list of concrete serialized styles.

Presentation may come from the placement paragraph, a list/caption/prepared-table prototype, a style/numbering resource, or placeholder run formatting. Configure a concrete usable numbered style in strict mode, for example:

```yaml
styles:
  lists:
    unordered: "List Bullet 2" # must have actual bullet numbering
    ordered: "List Number"     # must have actual decimal numbering
```

The UI name and internal ID may differ. `inspect styles` shows both. A name alone does not prove a numbering definition is usable.

If strict `check` reports an unavailable style/numbering, inspect the template and correct its resources. During exploration, `yadg check --yolo` and `yadg build --yolo` may choose a compatible resource or the real-numbering built-ins `builtin:unordered-list-v1` / `builtin:ordered-list-v1`. The fallback uses Word numbering definitions, never literal bullet/number glyphs. Each use is a visible degradation; it does not replace fixing the production template.

Use `yadg render --renderer libreoffice --renderer-path <bad-path> --yolo` to exercise the documented Word fallback (where Word is installed). The result reports requested versus actual renderer/runtime. If both renderers fail, no finalized output is claimed.
