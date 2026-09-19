# YADG Product Specification

## Purpose

YADG makes mandatory Word-based documentation maintainable as version-controlled Markdown without losing the structure and presentation supplied by prepared Microsoft Word templates.

The product is template-first rather than a general-purpose DOCX layout generator.

## Product authority model

The following ownership rules are normative:

1. Markdown is authoritative for maintainable document content and semantic document objects.
2. Prepared DOCX templates are authoritative for document structure, presentation, presentation vocabulary, numbering conventions, and Word-native document structures unless a template location explicitly delegates structure or placement to YADG.
3. Generated DOCX and PDF files are derived artifacts.
4. The Office-independent authoring pipeline is authoritative for semantic resolution and OOXML transformation.
5. Layout-dependent field evaluation and finalization belong to a separate renderer/finalizer.

YADG preserves template-owned formatting and Word-native conventions rather than recreating them from project-global configuration.

## Processing model

YADG has three conceptual stages:

1. parse/analyze Markdown into an OOXML-free semantic model;
2. author DOCX by resolving tags, template-local bindings, semantic references, and OOXML structures;
3. render/finalize layout-dependent document state in a real rendering engine.

Stages 1 and 2 do not require Microsoft Office.

## Workspace and CLI

A workspace is rooted at `YADG.md`.

The Office-independent CLI remains:

```text
yadg check [--workspace <path>]
yadg build [--workspace <path>]
```

Templates remain top-level `YadgTemplates/*.docx`; authored outputs remain same-named `YadgPreWords/*.docx`.

## Semantic model and stable identity

The semantic model remains OOXML-independent and includes headings, paragraphs, lists, tables, figures, structured-object anchors, supported inlines, and M0004 semantic cross-reference inlines.

Stable IDs are case-sensitive and workspace-wide:

```text
[A-Za-z][A-Za-z0-9_-]*
```

Display text, labels, captions, filenames, source positions, Word bookmark names, and field codes are not semantic identity.

## Sections

Section/content selection remains:

```text
{{content:<id>}}
{{section:<id>}}
```

A semantic section is numerically referenceable only when its Markdown heading is rendered exactly once and has effective Word numbering in that template. Detailed rules are in `docs/specs/WORD-REFERENCES.md`.

## Template vocabulary

Ordinary authoring tags remain:

```text
{{content:<id>}}
{{section:<id>}}
{{table:<id>}}
{{figure:<id>}}
{{value:<id>}}
```

`value` remains reserved and unsupported.

Template-control and prototype markers are defined separately in `docs/specs/WORD-REFERENCES.md`.

## Markdown support

M0004 adds semantic numeric references:

```markdown
See Figure [@system-context].
See Table [@interface-matrix].
See Section [@architecture].
```

`[@id]` emits only the numeric target result. Human-facing labels and grammar remain author/template-owned.

Ordinary Markdown hyperlinks remain unsupported.

## Structured content

Lists, generated tables, figures, assets, captions, and float-like placement remain governed by:

```text
docs/specs/STRUCTURED-CONTENT.md
```

M0004 extends table caption metadata and numbered caption/reference behavior through `docs/specs/WORD-REFERENCES.md`.

## Template-local front matter and prototypes

A template may contain visible removable YADG front matter that maps semantic roles to existing Word style IDs and template-owned prototypes.

Front matter is template-local, not workspace-global presentation configuration.

Caption prototypes may own literal labels, `SEQ` identifiers, field switches, punctuation, style, and formatting. YADG clones/adapts those structures instead of synthesizing localized presentation conventions.

The control format and compatibility rules are authoritative in `docs/specs/WORD-REFERENCES.md`.

## Cross-references

M0004 semantic references resolve by stable ID and are type-aware.

- figure/table references use bookmarked cloned `SEQ` fields plus Word `REF`;
- section references use uniquely rendered numbered headings plus Word `REF` paragraph-number semantics.

Generated bookmark names are implementation-private artifact mechanics.

## Template-owned indexes and lists

Existing template TOC, list-of-figures, and list-of-tables field structures remain template authority.

M0004 preserves them and authors compatible field structures, but does not evaluate or refresh them.

## Check and build

`check` validates the complete workspace and each template without producing authored outputs.

M0004 additionally validates:

- front-matter syntax/schema;
- mapped style existence;
- prototype existence/structure;
- semantic cross-reference resolution;
- target numbering eligibility and uniqueness;
- bookmark/field construction prerequisites.

`build` validates before modifying normal outputs.

M0004 may author bookmarks, `SEQ`, and `REF` structures, but displayed/cached field results are not authoritative until a renderer evaluates them.

## Renderer boundary

The later renderer/finalizer owns evaluation of `SEQ`, `REF`, TOC, list-of-figures, list-of-tables, pagination-dependent, and related Word field state.

## Simple values

`value` remains reserved; scalar value-source/override semantics remain undefined.

## Diagnostics

Diagnostics remain first-class and should identify stable code, source/template location, offending object/tag/binding, and corrective context where feasible.

## Non-goals

YADG is not a general-purpose Word layout engine or generic Markdown-to-DOCX converter, does not require Office for ordinary authoring, and does not treat generated artifacts as source authority.
