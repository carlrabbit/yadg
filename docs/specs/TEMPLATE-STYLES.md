# Template Styles and List Presentation Specification

## Status

Authoritative for YADG 1.1 template style discovery, style selector resolution, inline-code style binding, and list presentation.

Where this specification conflicts with earlier statements that style configuration values must be raw Word style IDs, this specification supersedes those statements.

Where this specification conflicts with the M0003 list requirement that list presentation must resolve exclusively through `ListBullet` / `ListNumber` paragraph styles, this specification supersedes that requirement.

The template remains authoritative for presentation.

## Why style discovery is required

WordprocessingML distinguishes:

- internal style ID;
- primary style/UI name;
- aliases;
- style type;
- visibility/UI flags;
- numbering definitions and style-to-numbering relationships.

Ordinary Word users primarily see names in the Word UI and should not be required to know internal IDs.

YADG therefore treats style IDs as an implementation identifier and user-visible names/aliases as first-class configuration selectors.

## Inspect styles

YADG 1.1 adds:

```text
yadg inspect styles [--workspace <path>] [--template <filename.docx>]
```

Without `--template`, inspect all top-level `YadgTemplates/*.docx`.

`--template` selects one top-level workspace template by filename. It must not escape `YadgTemplates`.

The command inspects the DOCX directly and does not require Markdown parsing, producer execution, or a valid `check`.

For every concrete style definition, output enough information to answer what a template author can configure:

- style type: paragraph, character, table, numbering;
- primary Word UI/display name;
- internal style ID;
- aliases;
- relevant visibility flags (`hidden`, `semiHidden`, `unhideWhenUsed`, `quickFormat`) when present;
- for paragraph styles, numbering capability/source when determinable.

Human-readable example shape:

```text
paragraph  "List Bullet 2"   id=ListBullet2  aliases=ListBullet  numbering=style
character  "Code"            id=Code          aliases=-           numbering=-
table      "Table Grid"      id=TableGrid     aliases=-           numbering=-
```

Exact spacing is not a compatibility contract.

Output is deterministic and grouped by template when multiple templates are inspected.

The command lists concrete serialized styles. Documentation explains that Word may expose latent/built-in gallery styles that are not serialized as usable concrete definitions until the document actually uses/materializes them.

## Style selector resolution

Existing template front-matter style values remain scalar strings.

A configured style string is resolved for the semantic role's required style type.

Resolution rules, in order:

1. exact ordinal match of internal style ID;
2. otherwise case-insensitive exact match of primary style name;
3. otherwise case-insensitive exact match of one alias;
4. no match -> configuration/validation error;
5. more than one candidate at a name/alias step -> ambiguity error listing candidate names and IDs.

An exact internal ID match wins before UI-name/alias matching to preserve 1.0 templates.

Aliases are read from the Word style definition using WordprocessingML alias semantics; YADG does not parse UI display strings heuristically.

Diagnostics for failed/ambiguous selectors tell the user to run `yadg inspect styles`.

Style selectors do not use substring/fuzzy matching.

## Role type constraints

Template roles require:

```text
heading1..heading9 -> paragraph
unordered           -> paragraph
ordered             -> paragraph
caption             -> paragraph
generatedTable      -> table
codeInline          -> character
```

A selector that uniquely resolves to the wrong style type is an error with the actual/required type shown.

The optional `codeInline` role is introduced by M0012.

There is no default `codeInline` style. It is required only when workspace `markdown.codeInline: style` applies to content rendered by that template.

## Front-matter compatibility

Template front matter remains schema version 1.

The existing style/prototype syntax is additively extended.

Conceptual example:

```yaml
version: 1

styles:
  headings:
    1: "Heading 1"
    2: "Heading 2"
  lists:
    unordered: "List Bullet 2"
    ordered: "List Number"
  generatedTable: "Table Grid"
  caption: "Caption"
  codeInline: "Code"

prototypes:
  figureCaption: figure-caption
  tableCaption: table-caption
  unorderedListItem: bullet-item
  orderedListItem: number-item
```

Existing raw IDs such as `Heading1`, `ListBullet`, and `TableGrid` remain valid when those IDs exist.

## Paragraph-style numbering

For style-based list rendering, YADG recognizes Word numbering semantics rather than requiring one narrow XML shape.

A paragraph style is considered list-capable when its effective Word numbering can be established from the concrete document through supported WordprocessingML relationships, including:

- direct/inherited paragraph-style numbering properties referencing a valid numbering instance;
- a valid numbering-level association to the paragraph style through the numbering definitions.

The implementation must not assume that a list is invalid merely because the paragraph style itself lacks a direct `numPr`.

If numbering resolution is absent or ambiguous for a selected style, `check` explains the condition and points to either another inspectable style or the list-item prototype escape hatch.

YADG does not synthesize an arbitrary new numbering definition merely to make an invalid style work.

## List item prototype escape hatch

A template may avoid style-to-numbering discovery entirely by binding:

```text
prototypes.unorderedListItem
prototypes.orderedListItem
```

to template-local prototypes.

When a prototype is configured for a list kind, that prototype is authoritative and takes precedence over the corresponding `styles.lists` binding for emitted list items.

### Prototype shape

A list-item prototype:

- contains exactly one Word paragraph;
- contains exactly one logical `{{item}}` placeholder;
- contains no other YADG authoring/control placeholder;
- has effective list numbering through its own paragraph/style/numbering structure;
- may use any template-owned paragraph/run formatting.

The prototype should be created in Word/LibreOffice as an actual list paragraph, not by editing raw OOXML.

Example conceptual visible content:

```text
• {{item}}
```

The visible bullet/number is produced by Word numbering, not literal bullet text in the placeholder contract.

### Authoring behavior

For every Markdown list item, YADG clones the prototype paragraph and substitutes `{{item}}` with the semantic inline content.

The clone preserves the prototype's:

- paragraph style;
- numbering properties;
- indentation;
- paragraph formatting;
- unrelated run formatting.

The placeholder's run formatting is the base formatting for ordinary item text; Markdown strong/emphasis/code/reference semantics overlay as applicable.

Configured ordered-list prototype clones must remain part of the same effective numbering sequence as intended by the prototype/template.

The prototype itself and control markers are removed from authored output.

### Validation

When a list prototype is selected, validation verifies the prototype rather than requiring the list style role to independently resolve numbering.

If prototype numbering is not valid, diagnostics identify the prototype and structural reason.

## List compatibility

If no list prototype is configured, the existing style-based approach remains and now uses the improved style selector/numbering rules.

This preserves valid 1.0 templates while making Word UI names and aliases usable.

## Inline code rendering

When workspace configuration selects:

```yaml
markdown:
  codeInline: style
```

YADG renders the code span using the template's `codeInline` character style.

Only the code span receives that character style.

The surrounding paragraph retains its normal paragraph/run presentation.

Strong/emphasis nesting with inline code must have deterministic semantics; implementation must preserve the code character style while applying supported explicit inline emphasis where both occur.

`markdown.codeInline: ignore` does not require the template role and renders code content as ordinary text.

## Documentation requirements

The user templates guide must include a troubleshooting section for:

```text
List style '<x>' does not resolve to a valid numbering definition.
```

It must explain:

- style ID vs UI name vs aliases;
- how to run `yadg inspect styles`;
- why renaming a style in Word may not change its internal ID;
- why Word may show a gallery/list choice that is not a concrete usable paragraph style;
- how to configure the actual display name/alias;
- when to use a list-item prototype instead.

Users are not expected to inspect raw `styles.xml` or `numbering.xml`.

## Validation evidence

Focused tests cover style ID/name/alias selection, ambiguity, type mismatch, numbering relationship variants, hidden/semi-hidden styles, and list prototypes.

Authoritative realistic compatibility uses actual Word/LibreOffice-produced templates, including at least one Word-origin list scenario whose UI-visible style name/alias differs from its internal ID.

## Non-goals

M0012 does not:

- expose arbitrary raw numbering IDs as normal workspace configuration;
- create a Word style editor;
- create new Word styles on behalf of the user;
- support nested Markdown lists;
- make fuzzy style matching a public behavior;
- promise that every latent style visible in the Word gallery is directly bindable.

## Presentation and resilience authority

Styles are one mechanism under the common template-example model. `docs/specs/TEMPLATE-PRESENTATION.md` governs strict/YOLO role resolution and deterministic related-resource selection. `docs/specs/AUTHORING-RESILIENCE.md` governs explicit recoveries, visible degradation diagnostics, renderer substitution, and fatal boundaries. Prior strict style/list requirements in this document continue to define strict mode; they do not prohibit the specifically authorized YOLO built-in examples.

Use `yadg inspect template` to see role resolution and safe fallback previews. `inspect styles` remains the concrete style inventory.
