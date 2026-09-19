# Architecture

## Architectural objective

YADG separates content semantics, Word package transformation, template-local Word conventions, and document rendering so each dependency is owned by the layer that requires it.

## Durable components

### `authoring-core`

Responsibilities:

- discover/parse Markdown;
- build the OOXML-free semantic model;
- resolve stable IDs and semantic cross-references;
- perform content/workspace validation independent of Word rendering.

Must not depend on Office, Office Interop, Windows-only APIs, or OOXML element types.

### `word-authoring`

Responsibilities:

- inspect prepared DOCX templates;
- discover visible YADG authoring tags and template-control regions;
- parse template-local front matter;
- map semantic roles to existing Word styles/prototypes;
- validate and clone template-owned prototypes;
- create Word-native bookmarks and field structures required by semantic references;
- preserve unrelated template structure/styles/fields;
- produce structurally complete DOCX artifacts.

May use Open XML SDK; must not require Word.

### `cli`

Owns deterministic commands, workspace selection, orchestration, exit codes, and diagnostics. Business/document semantics do not belong in argument plumbing.

### `renderer`

Owns the engine-specific finalization boundary. M0005 provides an isolated LibreOffice/UNO renderer that opens authored DOCX packages, refreshes fields and indexes, saves the finalized DOCX, and exports a matching PDF. It reports the exact executable, version, session profile, and output provenance. The renderer does not change semantic Markdown or authoring-core types.

Future rendering engines may be added behind this boundary, but field and pagination equivalence must be validated per engine. Microsoft Word and Office Interop are not dependencies of the M0005 authoring or validation graph.

## Artifact pipeline

```text
Markdown
  -> semantic model + semantic references
  -> Office-independent validation

DOCX template
  + authoring tags
  + optional visible front matter
  + optional Word-native prototypes
  -> OOXML authoring
  -> bookmarks + SEQ/REF fields
  -> authored DOCX
     (field results not authoritative)
  -> renderer/finalizer
  -> finalized DOCX / later PDF
```

## Template control region

Template metadata remains visible ordinary Word content.

When present, the removable initial control region contains front matter and prototypes. It is authoring metadata, not output content.

Exact syntax is in `docs/specs/WORD-REFERENCES.md`.

## Prototype ownership

A prototype is a template-owned Word structure that YADG clones/adapts.

For M0004 caption prototypes own label text, `SEQ` field identifier/switches, punctuation, style, and formatting. YADG owns semantic substitution and reference-target adaptation only.

## Identity boundary

Semantic stable IDs and Word bookmark names are separate namespaces.

Stable IDs belong to the semantic model. Generated bookmark names belong to the authored artifact and are implementation mechanics.

## Rendering boundary

OOXML authoring may create structurally valid fields but does not calculate Word field results.

Field evaluation, pagination, and TOC/list refresh belong to renderer/finalizer validation.

## Style ownership

Existing template styles remain authoritative.

M0004 front matter may map semantic roles to alternative existing style IDs; it does not define formatting properties.

## Alternative renderers

Alternative renderers remain possible later, but semantic equivalence of field evaluation is not assumed without validation.

## Security and path handling

Existing workspace path/link constraints remain in force. M0004 introduces no network/include boundary.
