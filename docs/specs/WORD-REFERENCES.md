# Word References and Template Metadata Specification

## Status

Authoritative for M0004 template front matter, template prototypes, numbered captions, bookmarks, and semantic cross-references.

## Design principle

Presentation vocabulary and Word-native conventions are template-owned. YADG front matter binds semantic roles to existing template structures; it does not define labels, fonts, numbering formats, or punctuation.

## Compatibility

Front matter is optional. Templates without it retain M0003 defaults:

```text
Heading1 ... Heading6
ListBullet
ListNumber
TableGrid
Caption
```

and retain M0003 unnumbered caption behavior.

## Template control region

If present, the control region starts at the first non-empty main-body paragraph:

```text
{{yadg:frontmatter}}
```

and front matter ends at:

```text
{{/yadg:frontmatter}}
```

Each intervening Word paragraph contributes one YAML line using logical text; blank paragraphs contribute blank lines.

Zero or more prototype blocks may immediately follow. Normal document content must not be interleaved with control content.

All control markers, YAML paragraphs, prototype markers, and prototype content are removed from authored output.

Control content is visible ordinary Word text; content controls or hidden custom XML are not required.

## Front matter schema version 1

Example:

```yaml
version: 1

styles:
  headings:
    1: Heading1
    2: Heading2
    3: Heading3
    4: Heading4
    5: Heading5
    6: Heading6
  lists:
    unordered: ListBullet
    ordered: ListNumber
  generatedTable: TableGrid
  caption: Caption

prototypes:
  figureCaption: figure-caption
  tableCaption: table-caption
```

Unknown keys are errors.

Style values are existing Word style IDs. Omitted bindings fall back to M0003 defaults. A configured style changes only which existing template style represents a semantic role.

`prototypes.figureCaption` and `prototypes.tableCaption` refer to template-local prototype IDs matching:

```text
[A-Za-z][A-Za-z0-9_-]*
```

Prototype IDs are not workspace semantic IDs.

## Prototype blocks

Prototype blocks use marker paragraphs:

```text
{{yadg:prototype:<prototype-id>}}
```

and:

```text
{{/yadg:prototype:<prototype-id>}}
```

M0004 caption prototypes contain exactly one Word paragraph between the markers.

A valid caption prototype contains:

1. exactly one Word `SEQ` field, simple or complex;
2. exactly one logical textual placeholder `{{caption}}`;
3. no other YADG authoring/control placeholder.

The prototype owns literal label text, sequence identifier, field switches, punctuation, paragraph/run formatting, and style.

Conceptually:

```text
Abbildung { SEQ Abbildung }: {{caption}}
```

is valid, but the actual Word field structure is authoritative.

When authoring, YADG clones the prototype, replaces `{{caption}}` with semantic caption text, preserves unrelated prototype structure, and creates an internal bookmark around the cloned sequence target.

## Semantic captions

Figure alt text remains the semantic figure caption. Empty alt text means no caption.

If no figure-caption prototype is configured, M0003 plain-caption behavior remains and the figure is not a numeric target.

M0004 extends table attributes:

```markdown
{#interface-matrix caption="Interfaces"}
```

The M0003 form remains valid:

```markdown
{#interface-matrix}
```

The attribute line supports only the stable ID and optional `caption` attribute in M0004. Unknown attributes are errors.

A captioned table without a configured table-caption prototype may render plain caption text using the mapped/default caption style, but is not a numeric target.

## Semantic cross-reference syntax

M0004 introduces:

```markdown
[@<stable-id>]
```

Examples:

```markdown
See Figure [@system-context].
See Table [@interface-matrix].
See Section [@architecture].
```

The syntax emits only the numeric Word result. Surrounding Markdown supplies human-facing labels and grammar.

References are supported in existing inline-capable contexts: ordinary paragraphs, flat list items, and table cells.

Ordinary Markdown hyperlinks remain outside M0004.

## Figure/table targets

A figure/table reference is valid in a template only when:

- the ID resolves to that semantic object;
- the object renders exactly once in that template;
- it has non-empty semantic caption text;
- the matching caption prototype is configured and valid.

YADG bookmarks the cloned `SEQ` target and emits a Word `REF` field with hyperlink behavior.

If the target renders zero or multiple times, the numeric reference is invalid.

## Section targets

A section reference is valid only when the referenced Markdown heading renders exactly once in that template and has effective Word paragraph numbering.

YADG bookmarks the rendered heading and emits a Word `REF` field using paragraph-number plus hyperlink semantics.

`{{content:id}}` does not render the referenced heading and therefore does not by itself create a section-number target.

M0004 does not bind Markdown section IDs to separately template-owned heading paragraphs. That requires a later explicit contract.

## Internal bookmarks

Generated Word bookmark names are artifact mechanics, not public YADG syntax.

They must be valid and unique within the document and consistently referenced by generated fields. Stable IDs remain the only author-facing semantic identity.

## Field evaluation boundary

M0004 authors field structure but does not evaluate it.

After `build`, `SEQ`/`REF` cached results may be stale, placeholder, or inherited prototype results. Those values are not correctness evidence.

The later renderer owns evaluation of:

- `SEQ`;
- `REF`;
- TOC;
- list of figures;
- list of tables;
- pagination-dependent and related Word fields.

## Template-owned TOC and lists

M0004 does not generate TOC/list-of-figures/list-of-tables from scratch.

Existing template field structures are preserved. Caption prototype sequence identifiers remain template authority and may align with those existing list fields.

## Validation

`check` validates real DOCX structure and diagnoses at least:

- malformed/duplicate front matter;
- unsupported schema version;
- unknown YAML keys;
- nonexistent mapped styles when used;
- missing/duplicate prototype IDs;
- invalid caption prototype `SEQ`/`{{caption}}` counts;
- unresolved/type-invalid semantic references;
- uncaptioned figure/table numeric references;
- missing required prototype;
- non-unique rendered targets;
- unnumbered section targets.

## Deferred behavior

M0004 does not define:

- binding stable Markdown sections to separately template-owned headings;
- deliberate duplicate targets with disambiguation;
- auto-injected localized labels;
- page-number references;
- workspace-global presentation configuration;
- prepared-table row population;
- field evaluation/rendering;
- PDF output.
