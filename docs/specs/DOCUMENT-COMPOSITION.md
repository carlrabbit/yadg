# Document Composition and Template Compatibility Specification

## Status

Authoritative for realistic template composition behavior introduced by M0010.

## Purpose

YADG is template-first. A production template is not merely a container for one placeholder paragraph; it is an existing Word document with its own layout, heading hierarchy, stories, fields, styles, sections, headers, footers, shapes, notes, and other Word-owned structures.

M0010 establishes two related contracts:

1. Markdown headings and paragraphs compose correctly into an existing template outline rather than assuming the Markdown document starts at Word heading level 1.
2. Workspace values behave like visible template-text substitution across the supported Word document stories rather than being limited to the main body/header/footer implementation paths used by M0006.

M0010 also establishes realistic application-produced DOCX fixtures as authoritative compatibility proofs before V1 release-readiness work.

## Template ownership

The existing ownership model remains:

```text
Template owns Word structure and presentation.
Markdown owns maintainable semantic document content.
Workspace values populate visible template text.
```

M0010 does not introduce Markdown syntax for creating Word-specific stories such as footnotes, endnotes, comments, or text boxes.

Those structures remain template-owned.

## Relative heading composition

### Principle

Markdown heading levels describe hierarchy inside the selected Markdown section.

When a `section` or `content` selection is inserted into an existing Word template outline, emitted Word headings are rebased relative to the template heading context at the insertion anchor.

A Markdown heading is therefore not automatically emitted as the same absolute Word heading level.

### Template heading context

For a `{{section:<id>}}` or `{{content:<id>}}` placement paragraph in the main document body, YADG determines the template context level from the nearest preceding top-level body paragraph, excluding the placement paragraph itself, whose effective Word paragraph outline level is 1 through 9.

The effective outline level must be resolved from Word paragraph/style semantics, including inherited paragraph-style properties. It must not be inferred from a style name such as `Heading4`.

If no preceding template paragraph has an effective outline level, the template context level is:

```text
0
```

Generated headings inserted by another YADG placement do not retroactively define the context of a later placement. Context is derived from the prepared template before YADG block replacement.

Static non-heading template paragraphs between the governing template heading and the placement do not break the context.

### Selected Markdown root

Let:

```text
C = template context level, 0..9
R = source Markdown level of the selected section root, 1..6
L = source Markdown level of an emitted Markdown heading
```

#### `section`

`{{section:id}}` includes the selected root heading.

The selected root becomes the immediate child heading of the template context:

```text
effective level = C + 1 + (L - R)
```

Therefore:

```text
template Heading 4
{{section:x}}

# X
## Child
```

when `# X` is the selected root yields:

```text
effective Heading 5: X
effective Heading 6: Child
```

The same rule applies when the selected Markdown root itself is not level 1. A selected source level-3 heading inserted below template outline level 4 still becomes effective level 5; its source descendants retain their relative level deltas.

#### `content`

`{{content:id}}` omits the selected root heading because the template is providing the structural/container context.

For descendant headings:

```text
effective level = C + (L - R)
```

A direct Markdown child of the omitted root therefore becomes the immediate child heading of the template context.

Example:

```text
template Heading 4
{{content:x}}

# X
## Child
```

yields:

```text
effective Heading 5: Child
```

Ordinary body content of `X` is inserted below the template Heading 4 without synthesizing the omitted `X` heading.

### Level gaps

YADG preserves source hierarchy deltas.

A source jump from the selected root at level 1 directly to a descendant level 3 remains a two-level jump after rebasing.

YADG does not silently normalize skipped heading levels.

### Supported effective levels

Effective Word heading levels are:

```text
1..9
```

If rebasing requires an effective level greater than 9, `check` and `build` fail before output mutation.

An effective level less than 1 is impossible under the defined formulas.

### Template heading style bindings

Template front matter version remains 1.

Heading style bindings are extended from levels `1..6` to:

```text
1..9
```

Default role/style mapping becomes:

```text
heading1 -> Heading1
...
heading9 -> Heading9
```

The template may bind any required effective level to another existing Word paragraph style.

A required effective heading style must exist.

The effective heading style, not the source Markdown heading level, controls Word numbering/outline behavior and numeric section-reference eligibility.

### Section references and indexes

When a rendered section heading is a numeric reference target, validation uses the effective heading level/style after rebasing.

Template-owned TOC/index behavior remains renderer-owned, but the authored document must expose the rebased heading structure so Word/LibreOffice can update those structures correctly.

A `content` selection still does not render its selected root heading and therefore does not create a numbered target for that omitted root merely because the surrounding template has a heading.

## Paragraph composition

Each Markdown paragraph remains a distinct Word paragraph.

For ordinary Markdown paragraphs emitted through a `section`/`content` placement, the placement paragraph remains the template paragraph prototype for ordinary paragraph properties/style, as in the existing authoring model.

M0010 explicitly requires realistic proof that:

- multiple consecutive Markdown paragraphs remain distinct paragraphs;
- the template-owned body style/paragraph properties survive;
- emphasis, strong emphasis, hard/soft breaks, and references remain correctly authored;
- heading rebasing does not accidentally apply heading formatting to ordinary paragraphs;
- static template paragraphs before and after the placement remain unchanged.

Lists, generated/prepared tables, figures, captions, and their existing style/prototype contracts remain unchanged.

## Supported visible-text value stories

`{{value:<id>}}` is a visible template-text substitution mechanism.

M0010 supports ordinary visible Word text in the following story/part classes:

- main document body, including table cells;
- all header parts, including default/first/even variants and their table cells;
- all footer parts, including default/first/even variants and their table cells;
- footnotes;
- endnotes;
- comments;
- text boxes / shape text represented through ordinary Wordprocessing text-box content in any of the supported parts above.

Existing visible text wrapped in hyperlinks or content controls is not excluded merely because of the wrapper, provided the logical tag is ordinary visible text in one supported replacement scope.

### Replacement scope

A value tag may be split across Word runs/text nodes, as before.

The logical replacement scope is one ordinary Word paragraph within one supported story/text container.

A tag must not cross:

- paragraph boundaries;
- part/story boundaries;
- separate nested text-box paragraph boundaries.

Nested text-box paragraphs are independent replacement scopes and must not be concatenated into the containing outer paragraph's logical text.

### Formatting

Replacement remains literal and non-recursive.

Replacement text inherits formatting from the run/text position containing the first logical character of the tag, preserving the M0006 rule.

Surrounding runs, paragraph formatting, fields, relationships, drawing geometry, note/comment identity, and story relationships remain template-owned.

### Validation

`check` and `build` enumerate the same supported value stories.

Malformed value tags, missing values, or otherwise invalid supported-story value substitutions fail validation before normal output mutation.

A shared header/footer part referenced by multiple Word sections is validated/replaced as one part; YADG does not create duplicate parts merely because multiple sections reference it.

## Value exclusions

M0010 does not treat the following as visible template text replacement targets:

- field instruction code (`w:instrText` or equivalent instruction attributes);
- relationship targets;
- bookmark names;
- document/custom properties;
- custom XML;
- package metadata;
- drawing/image alternative-text attributes;
- arbitrary XML attributes;
- YADG front-matter/prototype control syntax.

Tracked-change authoring semantics are not introduced by M0010. Compatibility fixtures must have tracked changes accepted before they become fixture authority.

M0010 does not add special semantics for a value deliberately placed inside a field result that a renderer later recalculates; such a template is not part of the compatibility proof.

## Realistic compatibility fixtures

M0010 requires two committed, redistribution-safe DOCX template fixtures:

1. one originating as a new document in desktop Microsoft Word;
2. one originating as a new document in LibreOffice Writer.

The fixtures must not be synthesized with Open XML SDK and then merely resaved for provenance.

Their required document structures/content must be created through the originating application's document model/UI and saved as DOCX by that application.

Tests use the committed binary fixture bytes as inputs and must not regenerate them on each run.

### Provenance

Repository-local fixture provenance records, for each fixture:

- origin application;
- concrete application version;
- platform;
- creation/save method sufficient to establish application origin;
- SHA-256 of the committed fixture;
- creation/update date;
- synthetic/non-confidential statement.

If a fixture is materially regenerated, provenance and hash must be updated deliberately.

### Required realistic content

Both fixtures must contain realistic template-owned layout and Word structures sufficient to prove the M0010 contract.

Across each fixture, include at least:

- non-default page/layout properties;
- template-owned static paragraphs before/between/after insertion locations;
- an existing template heading hierarchy with a placement below effective outline level 4;
- body-text placement paragraphs with nontrivial paragraph/style formatting;
- default header and footer with surrounding static text and existing fields such as page numbering;
- a footnote containing a value tag;
- an endnote containing a value tag;
- a comment containing a value tag;
- a visible text box/shape containing a value tag;
- value tags split across runs in at least one non-main-body story;
- ordinary main-body table/template content;
- existing Word-native fields/index structures appropriate to the fixture;
- at least one existing figure/table/layout element whose preservation can be inspected.

At least one of the two fixtures additionally exercises first-page or even-page header/footer behavior and more than one Word section/page-layout region.

The Word-origin and LibreOffice-origin fixtures may encode equivalent structures differently. That difference is part of the compatibility proof.

## Compatibility matrix

Authoring validation/build must run against both committed fixtures.

Tier-3 runtime validation covers all combinations:

```text
Word-origin template        -> build -> Word renderer
Word-origin template        -> build -> LibreOffice renderer
LibreOffice-origin template -> build -> Word renderer
LibreOffice-origin template -> build -> LibreOffice renderer
```

All four paths must complete successfully for the fixture scenarios.

M0010 does not claim byte-identical or pagination-identical output across renderers.

The matrix proves that realistic application-produced DOCX templates survive YADG authoring and can be finalized by either supported renderer.

## Non-goals

M0010 does not add Markdown syntax or semantic objects for:

- footnotes;
- endnotes;
- comments;
- text boxes/shapes;
- headers;
- footers.

M0010 does not make block YADG placement tags valid in those stories.

`section`, `content`, `table`, `figure`, and prepared-table placement remain governed by their existing main-document-body contracts.

M0010 does not add generic support for every OOXML part, tracked-change authoring, arbitrary XML text substitution, new image formats, new Markdown block types, NuGet packaging, GitHub workflows, or V1 release publication automation.
