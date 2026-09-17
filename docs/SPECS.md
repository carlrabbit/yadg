# YADG Product Specification

## Purpose

YADG exists to make mandatory Word-based documentation maintainable as version-controlled Markdown without losing the exact structure and presentation supplied by institutional Word templates.

The product is intentionally template-first rather than a general-purpose DOCX layout generator.

## Product authority model

The following ownership rules are normative:

1. Markdown is authoritative for maintainable document content.
2. Prepared DOCX templates are authoritative for document structure and presentation unless a template location explicitly delegates structure to Markdown.
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

### Root and invocation

A YADG workspace is a directory containing a root-level file named exactly:

```text
YADG.md
```

M0002 processes exactly one workspace per `check` or `build` invocation.

The CLI contract is:

```text
yadg check [--workspace <path>]
yadg build [--workspace <path>]
```

`--workspace` defaults to the current working directory.

A supplied workspace path must identify the workspace root itself. Parent-directory discovery of multiple workspaces is not part of M0002.

The M0001 prototype arguments `--markdown`, `--template`, and `--output` are not a compatibility surface and may be removed when the workspace contract is implemented.

### Workspace marker

`YADG.md` identifies the workspace but is not document source content.

For M0002, its body may contain human-readable notes. M0002 does not define machine-readable configuration inside `YADG.md`; future configuration must be introduced by later project authority rather than inferred from arbitrary Markdown content.

### Conventional directories

M0002 uses these fixed workspace conventions:

```text
<workspace>/
  YADG.md
  YadgTemplates/
    *.docx
  YadgPreWords/
    <generated outputs>
  ... Markdown sources ...
```

`YadgTemplates/` contains prepared source templates.

`YadgPreWords/` contains authored DOCX outputs and is derived-artifact storage.

For every top-level `.docx` file directly inside `YadgTemplates/`, a successful build produces or replaces:

```text
YadgPreWords/<same-file-name>.docx
```

M0002 does not recursively discover templates below `YadgTemplates/`.

A valid workspace contains at least one template and at least one Markdown source file.

### Markdown source discovery

Markdown source files are `.md` files discovered recursively below the workspace root, excluding:

- the root `YADG.md`;
- `YadgTemplates/`;
- `YadgPreWords/`;
- directories whose names begin with `.`.

Source-file order has no document-composition meaning. Templates select semantic objects through stable IDs.

A second `YADG.md` discovered below the workspace root is a nested-workspace error. Nested workspaces are not implicitly processed or shared.

M0002 does not implement Markdown link/include files or any other cross-workspace sharing mechanism.

Workspace discovery must not follow symbolic-link/reparse-point directories. A candidate Markdown source or template that is itself a symbolic link/reparse point is invalid for M0002. Discovered product inputs therefore remain physically within the declared workspace.

## Markdown semantic model

The Markdown parser exposes structured semantic nodes rather than pre-rendered text.

M0002 requires semantic representation sufficient for:

- headings;
- paragraphs;
- inline text;
- emphasis;
- strong emphasis;
- soft line breaks;
- hard line breaks;
- stable section IDs;
- document/section hierarchy.

OOXML/Open XML SDK types must not appear in the semantic Markdown model.

### M0002 supported Markdown subset

M0002 authors:

- ATX headings (`#` through `######`);
- paragraphs;
- plain inline text;
- emphasis;
- strong emphasis;
- soft line breaks, rendered as normal paragraph whitespace;
- Markdown hard line breaks, rendered as Word line breaks.

A heading may end with an explicit stable ID:

```markdown
## Architecture {#architecture}
```

The stable-ID suffix is metadata and is not part of rendered heading text.

Headings without IDs remain semantic headings and may appear inside selected content, but only an object with a stable ID can be referenced directly from a template.

M0002 does not author lists, tables, images, block quotes, fenced/indented code blocks, raw HTML, or Markdown hyperlinks. Encountering an unsupported construct in a discovered source is a `check` error; it must not be silently flattened to plain text.

Later milestones may extend the supported Markdown subset without changing the reference identity model.

## Stable reference semantics

Stable IDs are case-sensitive and workspace-wide.

The M0002 stable-ID syntax is:

```text
[A-Za-z][A-Za-z0-9_-]*
```

Duplicate stable IDs anywhere in one workspace are errors.

Changing a heading caption must not require changing a template reference when the heading's stable ID remains unchanged.

Heading-based selection distinguishes:

- **section** — the referenced heading plus its complete section body;
- **content** — the referenced section body without the referenced heading.

Nested headings belong to the parent section body until a heading of the same or higher level terminates the section.

Example:

```markdown
## Architecture {#architecture}

Intro.

### Components

Component text.

### Deployment

Deployment text.
```

`{{section:architecture}}` selects the `## Architecture` heading and all shown content.

`{{content:architecture}}` selects the intro, `### Components`, `### Deployment`, and their bodies, but excludes the `## Architecture` heading.

The design must not model this distinction as generic `append` versus `replace`.

## Word template tags

### Vocabulary

The explicit textual tag vocabulary is:

```text
{{content:<stable-id>}}
{{section:<stable-id>}}
{{value:<stable-id>}}
```

Tag keywords and stable IDs are case-sensitive. Whitespace inside a tag is not permitted.

M0002 implements `content` and `section`.

`value` is reserved as the scalar-replacement vocabulary for a later milestone. A `value` tag encountered during M0002 is diagnosed as recognized-but-unsupported rather than ignored.

The M0001 prototype spelling:

```text
{{yadg:section:content:<stable-id>}}
```

is not a supported compatibility surface after M0002.

### Block-tag placement

`content` and `section` are block tags.

For M0002 a block tag must be the only non-whitespace logical content of a Word paragraph in the main document body.

A block tag embedded in ordinary prose, or located in a table, header, footer, textbox, footnote/endnote, comment, or other unsupported Word part/container, is invalid for M0002 and must produce a diagnostic rather than being silently ignored.

YADG must detect YADG-shaped tags in unsupported locations sufficiently to report that the template is not supported instead of incorrectly declaring it valid.

### OOXML run handling

A visually contiguous textual tag can be split across multiple OOXML runs by Word.

Tag discovery therefore operates on logical paragraph text, not individual `w:t` elements.

Replacement must preserve unaffected template-owned structure and formatting.

## Word block rendering for M0002

The block-tag paragraph is an insertion anchor and is replaced by the selected Markdown block sequence.

### Paragraphs

Generated ordinary Markdown paragraphs inherit the insertion anchor's paragraph properties/style.

M0002 relies on template styles for presentation. It does not copy arbitrary font/paragraph settings from Markdown.

Arbitrary direct run formatting from the tag text itself is not a formatting contract; template authors should express paragraph presentation through Word styles.

### Headings

Generated Markdown headings retain their Markdown heading level.

They are authored using Word paragraph style IDs:

```text
Heading1
Heading2
Heading3
Heading4
Heading5
Heading6
```

according to the Markdown level.

The template remains authoritative for what those styles look like. If a selected heading requires a corresponding heading style that is absent from the template, `check` fails rather than synthesizing a replacement style.

### Inlines

Strong emphasis is represented semantically as bold Word run formatting.

Emphasis is represented semantically as italic Word run formatting.

Hard line breaks become Word line breaks within the paragraph.

No Markdown inline syntax is emitted literally merely because the renderer lacks semantic support.

## Template-owned headings

Where a Word template already owns a heading, the normal pattern is:

```text
Template heading
{{content:architecture}}
```

Where a template intentionally delegates the referenced heading and body to Markdown, use:

```text
{{section:architecture}}
```

This distinction is visible in the template and statically validated.

## Check semantics

`yadg check` validates the complete workspace without producing normal authored outputs.

M0002 validation includes at least:

- workspace marker and directory validity;
- nested workspace detection;
- regular-file/path-boundary rules;
- source and template presence;
- Markdown parsing and supported-subset validation;
- workspace-wide duplicate stable IDs;
- template tag syntax;
- unsupported tag kinds/locations;
- unresolved stable IDs;
- required Word heading styles for selected headings.

Validation should accumulate independent diagnostics where practical rather than stopping at the first unrelated error.

A successful `check` means the workspace is eligible for Office-independent authoring under the implemented feature set. It does not claim Microsoft Word rendering/layout correctness.

## Build semantics

`yadg build` first performs the same semantic/template validation required by `check`.

If validation fails, `build` fails before modifying normal output artifacts.

On successful validation, each top-level template in `YadgTemplates/` is copied/transformed into `YadgPreWords/` with the same filename.

Existing output files corresponding to current templates may be replaced.

Build does not delete unrelated files from `YadgPreWords/`.

An operational I/O failure after output writing begins fails the command and must be reported clearly; M0002 does not require transactionally rolling back already written artifacts after such an external failure.

A successfully completed build must not leave unresolved supported YADG tags in authored outputs.

## Tables and figures

Markdown tables and figures are outside M0002.

A later structured-content milestone will define their semantic nodes, template-population rules, captions, and Word-specific structures.

## Lists

Markdown lists are outside M0002.

Word list numbering is a structured Word concern and must be specified with its numbering/style contract before implementation. M0002 diagnoses lists as unsupported rather than flattening or approximating them with literal bullet/number characters.

## Simple values

Simple substitutions such as dates, document names, versions, and other scalar content will use the reserved `value` vocabulary.

Their value-source and override semantics are outside M0002.

## Extensions

Extensions must feed the semantic YADG model rather than directly manipulate OOXML by default.

Two extension shapes are anticipated:

- Markdown transformations that convert specialized Markdown into semantic YADG nodes;
- content providers that convert external structured inputs into semantic YADG nodes.

Direct OOXML extension points require an explicit later architectural decision.

Extensions are outside M0002.

## CLI surface

The intended command vocabulary remains:

```text
yadg check
yadg build
yadg render
yadg publish
```

M0002 establishes workspace-aware `check` and `build`.

`render` and `publish` remain later capabilities.

## Diagnostics

Diagnostics are first-class product behavior.

Where feasible, diagnostics identify:

- a stable diagnostic code;
- source/template location;
- the offending reference or tag;
- enough context for a maintainer to correct the problem.

Exact diagnostic code allocation is implementation-owned as long as established codes are not silently repurposed to mean incompatible errors.

## Non-goals

YADG is not:

- a general-purpose Word layout engine;
- a Markdown-to-DOCX converter that ignores a prepared template;
- a system that requires Office for ordinary authoring;
- a mechanism for treating generated DOCX files as editable source authority;
- a `dotnet-library` product profile or reusable library product merely because implementation uses .NET class-library projects.
