# YADG Product Specification

## Purpose

YADG makes mandatory Word-based documentation maintainable as version-controlled Markdown without losing the structure and presentation supplied by prepared Microsoft Word templates.

The product is template-first rather than a general-purpose DOCX layout generator.

## Product authority model

The following ownership rules are normative:

1. Markdown is authoritative for maintainable document content and semantic document objects.
2. Prepared DOCX templates are authoritative for document structure and presentation unless a template location explicitly delegates structure or placement to Markdown/YADG.
3. Generated DOCX and PDF files are derived artifacts.
4. The Office-independent authoring pipeline is authoritative for semantic resolution and OOXML transformation.
5. Layout-dependent finalization belongs to a separate renderer/finalizer.

YADG preserves template-owned formatting by default rather than recreating Word presentation from Markdown.

## Processing model

YADG has three conceptual stages:

1. **Parse and analyze** Markdown into a semantic document model.
2. **Author** DOCX artifacts by resolving template tags and transforming OOXML.
3. **Render/finalize** layout-dependent document state in a real rendering engine.

Stages 1 and 2 must not require Microsoft Office.

Stage 3 is isolated from authoring. The initial renderer may use Microsoft Word through Office Interop, but the authoring architecture must not make Word a transitive dependency.

## Workspaces

A YADG workspace is a directory containing a root-level file named exactly:

```text
YADG.md
```

One `check` or `build` invocation processes exactly one workspace.

The CLI contract is:

```text
yadg check [--workspace <path>]
yadg build [--workspace <path>]
```

`--workspace` defaults to the current working directory.

`YADG.md` identifies the workspace but is not document source content. Its body may contain human-readable notes. No machine-readable configuration syntax inside `YADG.md` is currently defined.

The conventional workspace directories are:

```text
<workspace>/
  YADG.md
  YadgTemplates/
    *.docx
  YadgPreWords/
    <generated outputs>
  ... Markdown sources and assets ...
```

For every top-level `.docx` file directly inside `YadgTemplates/`, a successful build produces or replaces `YadgPreWords/<same-file-name>.docx`.

Markdown source files are `.md` files discovered recursively below the workspace root, excluding the root `YADG.md`, `YadgTemplates/`, `YadgPreWords/`, and dot-prefixed directories.

A second `YADG.md` below the workspace root is a nested-workspace error.

Workspace discovery must not follow symbolic-link/reparse-point directories. Candidate Markdown sources, templates, and referenced assets that are symbolic links/reparse points are invalid.

## Semantic document model

The Markdown parser exposes structured semantic nodes rather than pre-rendered text.

The semantic model supports document blocks and inlines independently of OOXML. Open XML SDK types must not appear in the semantic Markdown model.

Current block concepts include:

- headings;
- paragraphs;
- flat ordered and unordered lists;
- generated tables;
- figures;
- natural anchors for movable structured objects.

Current inline concepts include:

- text;
- emphasis;
- strong emphasis;
- soft line breaks;
- hard line breaks.

## Stable identity and references

Stable IDs are explicit, case-sensitive, and workspace-wide.

The stable-ID syntax is:

```text
[A-Za-z][A-Za-z0-9_-]*
```

A stable ID identifies one semantic object. Duplicate IDs are invalid even when the colliding objects have different semantic types.

Display text, heading text, captions, filenames, and source position are not semantic identity.

Changing presentation text must not require changing template references when the stable ID remains unchanged.

## Sections

A heading may end with an explicit stable ID:

```markdown
## Architecture {#architecture}
```

Heading-based selection distinguishes:

- **section** — referenced heading plus its complete section body;
- **content** — referenced section body without the referenced heading.

Nested headings belong to the parent section body until a heading of the same or higher level terminates the section.

Template syntax is:

```text
{{content:architecture}}
{{section:architecture}}
```

The distinction is semantic and must not be modeled as generic `append` versus `replace`.

## Template tag vocabulary

The explicit textual vocabulary is:

```text
{{content:<stable-id>}}
{{section:<stable-id>}}
{{table:<stable-id>}}
{{figure:<stable-id>}}
{{value:<stable-id>}}
```

Tag keywords and stable IDs are case-sensitive. Whitespace inside a tag is not permitted.

`content`, `section`, `table`, and `figure` are implemented authoring operations.

`value` remains reserved for a later scalar-value milestone and must be diagnosed as recognized-but-unsupported.

All current tags are block tags. A block tag must be the only non-whitespace logical content of a Word paragraph in the supported main document body.

Tags embedded in prose or unsupported Word parts/containers are invalid and must be diagnosed rather than silently ignored.

Tag discovery operates on logical paragraph text and must continue to work when Word splits one textual tag across multiple OOXML runs.

## Markdown support

The core Markdown subset supports:

- ATX headings (`#` through `######`);
- paragraphs;
- plain text;
- emphasis;
- strong emphasis;
- soft line breaks;
- hard line breaks;
- flat unordered lists;
- flat ordered lists;
- pipe tables as defined in `docs/specs/STRUCTURED-CONTENT.md`;
- block figures as defined in `docs/specs/STRUCTURED-CONTENT.md`.

Unsupported Markdown must be diagnosed rather than flattened or silently discarded.

Currently unsupported constructs include:

- nested lists;
- multi-paragraph list items;
- task lists;
- block quotes;
- fenced/indented code blocks;
- raw HTML;
- ordinary Markdown hyperlinks;
- inline/non-figure images;
- table row/column spans.

## Word block rendering

### Paragraphs

Generated ordinary paragraphs inherit the insertion anchor's paragraph properties/style.

### Headings

Generated Markdown headings retain their Markdown level and use existing Word paragraph style IDs `Heading1` through `Heading6`.

If a selected heading requires a corresponding style that is absent, `check` fails rather than synthesizing a style.

### Inlines

Strong emphasis becomes bold Word run formatting.

Emphasis becomes italic Word run formatting.

Hard line breaks become Word line breaks.

Soft breaks render as normal paragraph whitespace.

### Structured content

Lists, generated tables, figures, captions, asset rules, and float-like placement are defined by `docs/specs/STRUCTURED-CONTENT.md`.

## Check semantics

`yadg check` validates the complete workspace without producing normal authored outputs.

Validation includes, as applicable:

- workspace marker/directory validity;
- nested workspace detection;
- path/link boundary rules;
- source/template/asset presence;
- supported Markdown semantics;
- workspace-wide stable-ID uniqueness and type-correct references;
- template tag syntax and placement;
- unresolved references;
- required Word styles/numbering contracts;
- structured-object placement uniqueness;
- image format/dimension constraints.

Validation should accumulate independent diagnostics where practical.

A successful `check` means the workspace is eligible for Office-independent authoring under the implemented feature set. It does not claim final Microsoft Word layout correctness.

## Build semantics

`yadg build` first performs the same semantic/template validation required by `check`.

If validation fails, `build` fails before modifying normal output artifacts.

On success, each top-level template in `YadgTemplates/` is copied/transformed to a same-named output in `YadgPreWords/`.

Build does not delete unrelated files in `YadgPreWords/`.

An operational I/O failure after writing begins fails the command and is reported clearly; transactional rollback of already written files is not required.

A successfully completed build must not leave unresolved supported YADG tags in authored outputs.

## Template-owned headings

Where a template owns a heading, the normal pattern is:

```text
Template heading
{{content:architecture}}
```

Where the template delegates the heading and body:

```text
{{section:architecture}}
```

## Tables and figures

Tables and figures are semantic objects with natural source anchors and optional template placement overrides.

Their detailed contract is authoritative in:

```text
docs/specs/STRUCTURED-CONTENT.md
```

Prepared-table row population is not currently supported.

## Simple values

Scalar substitutions such as dates, document names, and versions will use the reserved `value` vocabulary.

Their source and override semantics remain undefined and unsupported.

## Extensions

Extensions must feed the semantic YADG model rather than directly manipulating OOXML by default.

Anticipated extension shapes are Markdown transformations and external content providers that produce semantic nodes.

Direct OOXML extension points require a later explicit architectural decision.

## CLI surface

The intended command vocabulary remains:

```text
yadg check
yadg build
yadg render
yadg publish
```

`check` and `build` are workspace-aware and Office-independent.

`render` and `publish` remain later capabilities.

## Diagnostics

Diagnostics are first-class product behavior.

Where feasible, diagnostics identify:

- a stable diagnostic code;
- source/template/asset location;
- offending reference or tag;
- enough context for correction.

Established diagnostic codes must not be silently repurposed to incompatible meanings.

## Non-goals

YADG is not:

- a general-purpose Word layout engine;
- a Markdown-to-DOCX converter that ignores prepared templates;
- a system requiring Office for ordinary authoring;
- a mechanism for treating generated DOCX files as editable source authority;
- a `dotnet-library` product merely because internal class-library projects exist.
