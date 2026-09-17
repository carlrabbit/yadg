# Architecture

## Architectural objective

YADG separates content semantics, Word package transformation, and document rendering so each dependency is owned by the layer that actually requires it.

## Durable components

The initial architecture recognizes these semantic components. Names describe responsibilities; implementation project names may differ if the boundaries remain intact.

### `authoring-core`

Responsibilities:

- discover and parse YADG Markdown inputs;
- build the semantic YADG document model;
- resolve stable IDs and references;
- perform project/content validation independent of DOCX rendering;
- expose extension abstractions that produce semantic model content.

Must not depend on Microsoft Office, Office Interop, or Windows-only APIs.

### `word-authoring`

Responsibilities:

- inspect prepared DOCX templates;
- discover visible textual YADG tags across supported OOXML structures;
- map semantic content into template-owned Word structures;
- create new OOXML elements only when the contract delegates their creation to YADG;
- preserve unaffected template structure and styles;
- produce structurally complete DOCX artifacts;
- perform DOCX/OOXML-specific validation.

This component may use an OOXML/Open XML SDK but must not require Word to be installed.

### `cli`

Responsibilities:

- expose deterministic product commands;
- resolve working-directory/project input;
- coordinate parsing, checking, authoring, and later renderer invocation;
- provide useful exit codes and diagnostics.

Business/document semantics do not belong in CLI argument plumbing.

### `word-renderer`

Responsibilities:

- open an authored DOCX using Microsoft Word on Windows;
- update layout-dependent fields, references, indexes/TOCs, and related state;
- save the finalized DOCX;
- optionally export PDF when the invoking command requires it.

This component is the only initial component allowed to require Office Interop.

It must be separately invokable so Office-independent environments can complete authoring without it.

## Dependency direction

The intended dependency direction is:

```text
CLI
 |  |  +--------------------> renderer contract/client
 v
authoring orchestration
 |  v  v
authoring-core <---- word-authoring adapter
```

The semantic model belongs to the Office-independent side of the system. Word-specific types must not leak into core document semantics.

The renderer consumes an authored artifact; it must not become the place where Markdown parsing or ordinary template replacement logic lives.

## Artifact pipeline

```text
Markdown sources
      |
      v
semantic YADG model
      |
      +---- validation diagnostics
      |
DOCX template
      |
      v
OOXML authoring
      |
      +---- structural validation evidence
      |
      v
authored DOCX
      |
      +----------------------> usable build artifact without Office
      |
      v
Word renderer/finalizer
      |
      +---- rendering evidence
      |
      +--> finalized DOCX
      +--> PDF
```

## Template text parsing boundary

Visible tags are parsed from logical supported Word text containers rather than individual OOXML runs.

The implementation must account for Word splitting a tag across runs due to formatting, revision/edit history, spellchecking, or other document operations.

Tag replacement must not flatten unrelated formatting merely to simplify matching.

## Style ownership

Existing template styles are authoritative.

When generated semantic elements require Word elements that do not already exist in the template, mapping from semantic roles to Word styles must be explicit. YADG should refer to Word styles rather than reimplement font/paragraph formatting as ad hoc properties.

## Rendering boundary

OOXML describes document content but does not perform Word layout.

Anything requiring actual pagination or Word field calculation belongs to renderer/finalizer validation. The authoring pipeline may preserve or create field structures, but it must not claim layout correctness from package manipulation alone.

## Alternative renderers

The architecture intentionally permits later renderer implementations using another engine such as LibreOffice.

Alternative renderer support is not part of M0001 and semantic equivalence between renderers must not be assumed without validation.

## Security and path handling

Future include/link features can cross filesystem boundaries. Before such behavior is implemented, project authority must define:

- canonical path resolution;
- allowed workspace/root boundaries;
- symlink/reparse-point behavior;
- cycle handling;
- maximum indirection;
- diagnostics.

An implicit “a Markdown file containing only a path means include” convention is not currently project authority.
