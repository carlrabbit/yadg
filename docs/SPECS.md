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

YADG should preserve template-owned formatting by default rather than copying presentation settings from Markdown.

## Processing model

YADG has three conceptual stages:

1. **Parse and analyze** Markdown into a semantic document model.
2. **Author** DOCX artifacts by resolving template tags and transforming OOXML.
3. **Render/finalize** layout-dependent document state in a real rendering engine.

Stages 1 and 2 must not require Microsoft Office.

Stage 3 is isolated from authoring. The initial implementation may use Microsoft Word through Office Interop, but the authoring architecture must not make Word a transitive dependency.

## Markdown semantic model

The Markdown parser must expose semantic nodes sufficient to represent at least:

- headings/sections;
- paragraphs with supported inline content;
- lists;
- tables;
- images/figures;
- stable references/IDs;
- simple values or substitutions;
- extension-produced semantic content.

The Word transformation layer consumes this semantic model. Markdown parsing must not directly manipulate OOXML.

## Stable reference semantics

References must use stable semantic IDs rather than presentation text.

Changing a heading caption must not require changing a template reference when the heading's stable ID remains unchanged.

At minimum, heading-based content selection distinguishes:

- **section** — the referenced heading plus its complete section body;
- **content** — the referenced section body without the referenced heading.

Nested subsections belong to the parent section body until a heading of the same or higher level terminates the section.

Example:

```markdown
## Architecture {#architecture}

Intro.

### Components

...

### Deployment

...
```

A `section` reference to `architecture` resolves the `## Architecture` heading plus all shown content. A `content` reference resolves `Intro`, `### Components`, and `### Deployment` with their respective bodies, but not the `## Architecture` heading.

The design must not model this distinction as a generic `append` versus `replace` switch.

## Word template tags

YADG templates use visible textual tags. This is an intentional authoring requirement.

Template maintainers must be able to:

- type a tag with ordinary Word editing;
- see the tag in document context;
- copy and search tags;
- preserve and inspect native Word structure such as headings and TOCs without YADG-specific Word tooling.

YADG must not require Word content controls for ordinary template authoring.

The exact finalized tag grammar is milestone-controlled until implemented, but it must remain intentionally small and must not evolve into a general programming language.

### OOXML run handling

A visually contiguous textual tag can be split across multiple OOXML runs by Word. Tag discovery therefore operates on logical text at a supported container boundary, not on individual `w:t` elements.

Replacement must preserve unaffected template-owned formatting and structure.

## Template-owned headings

Where a Word template already owns a heading, the normal pattern is to insert only Markdown section content beneath that heading.

Where a template intentionally delegates a section's heading and body to Markdown, a complete section selection may be used.

The product should make this ownership distinction visible and statically checkable.

## Tables and figures

Markdown tables and figures may populate prepared template structures or be inserted as generated semantic elements.

When a template provides a specialized Word table, template formatting remains authoritative. YADG may populate its rows/cells according to the tag contract.

Captions and referenceable figure/table identity must use stable semantic IDs.

## Simple values

Simple substitutions such as dates, document names, versions, and other scalar content must be representable as named values.

Dynamic values must remain explicitly addressable so callers can override them where project workflow requires it.

## Workspaces

A YADG working directory contains project-local configuration and Markdown source.

Workspace discovery, link/include behavior, path-boundary rules, and the final configuration filename remain implementation decisions for a later milestone unless promoted into authority before then.

Nested workspace behavior must be explicit before implemented; it must not emerge accidentally from recursive file discovery.

## Extensions

Extensions must feed the semantic YADG model rather than directly manipulate OOXML by default.

Two extension shapes are anticipated:

- Markdown transformations that convert specialized Markdown into semantic YADG nodes;
- content providers that convert external structured inputs into semantic YADG nodes.

Direct OOXML extension points require an explicit later architectural decision.

## CLI surface

The intended command vocabulary is:

```text
yadg check
yadg build
yadg render
yadg publish
```

Semantics:

- `check` parses and validates project and template consistency without producing normal product artifacts;
- `build` produces structurally complete DOCX artifacts without requiring Office;
- `render` performs layout-dependent finalization through an available renderer;
- `publish` composes build/finalization and later publication-specific preparation as project authority defines.

The initial milestone does not have to implement all four commands.

## Diagnostics

Diagnostics are first-class product behavior.

Where feasible, diagnostics identify:

- a stable diagnostic code;
- source/template location;
- the offending reference or tag;
- enough context for a maintainer to correct the problem.

`check` is expected to grow into a major product capability, covering unresolved references, duplicate IDs, malformed directives/tags, invalid template structure, unsupported Markdown constructs, assets, and integration constraints as those features are implemented.

## Non-goals

YADG is not:

- a general-purpose Word layout engine;
- a Markdown-to-DOCX converter that ignores a prepared template;
- a system that requires Office for ordinary authoring;
- a mechanism for treating generated DOCX files as editable source authority;
- a `dotnet-library` product profile or reusable library product merely because implementation uses .NET class-library projects.
