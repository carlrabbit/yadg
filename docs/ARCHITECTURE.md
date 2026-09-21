# Architecture

## Objective

YADG separates:

```text
workspace/content semantics
-> external content production
-> Office-independent DOCX authoring
-> renderer-specific finalization
-> explicit publication
```

## Authoring core

Owns workspace/front-matter discovery, Markdown parsing, semantic objects, values, references, and producer-independent validation.

It must not depend on Microsoft Word COM automation.

Markdown source heading levels remain semantic source hierarchy. Template-relative Word heading levels are an authoring concern because they depend on the concrete insertion context of each DOCX template.

## Word authoring

Owns prepared DOCX inspection/transformation through Open XML.

It creates structurally complete DOCX artifacts without requiring an Office renderer.

M0010 makes two template-sensitive authoring concerns explicit:

1. resolve each `section`/`content` placement's existing template outline context and map emitted Markdown headings to effective Word heading levels;
2. discover supported Word text stories/containers and apply workspace value substitution consistently across them.

The semantic Markdown model remains OOXML-free and does not store template-specific effective heading levels.

## Heading composition boundary

Heading rebasing belongs to Word authoring, not Markdown parsing.

For one selected section, Word authoring combines:

```text
selected Markdown root level
+ source descendant level delta
+ prepared-template outline context
= effective Word heading level
```

The prepared template context is resolved before placement replacement.

Generated headings from one placement do not become planning/context input to another placement.

Effective Word heading role/style and numbering validation operate on the mapped level.

## Visible-text story boundary

Workspace values remain ordinary source-controlled scalar data.

Word authoring is responsible for locating ordinary visible text in the supported story parts:

```text
main document
headers
footers
footnotes
endnotes
comments
nested text boxes / shape text
```

Each paragraph/text-container scope is processed independently so nested text boxes are not accidentally concatenated into an outer paragraph's logical text.

This is template text substitution, not semantic Markdown authoring in those stories.

## External content producers

Remain separate-process source adapters.

Mermaid output becomes an ordinary semantic figure before Word authoring.

## Renderer boundary

Renderers consume `YadgPreWords/*.docx` and produce finalized workspace artifacts.

### LibreOffice

Existing isolated UNO renderer.

Produces finalized DOCX and its existing PDF side output.

### Microsoft Word

Windows Microsoft Word automation specialization.

Word-specific COM types/code remain isolated from authoring-core and Open XML authoring.

The full Word-enabled product is allowed to be Windows-build-specific.

## Realistic fixture boundary

Focused synthetic OpenXML tests remain useful for precise edge cases.

They are not sufficient evidence for V1 template compatibility.

M0010 adds committed binary DOCX fixtures whose origin is real Microsoft Word or LibreOffice Writer and whose application provenance/hash is recorded.

Compatibility tests copy those immutable fixture bytes into temporary workspaces; they do not recreate the template structure with OpenXML SDK during the test.

## Publication boundary

Publishing is downstream of rendering and has no Word/LibreOffice dependency.

```text
YadgWords/*.docx
       |
       | yadg publish
       v
explicit filesystem destination/*.docx
```

The publisher copies bytes; it does not open or modify Word documents.

## Dependency direction

Semantic/core and Open XML authoring must not gain renderer-specific COM dependencies.

Renderer-specific runtime code depends inward on the common renderer contract.

Publishing depends only on workspace/finalized-artifact conventions and filesystem operations.

## Platform specialization

Windows + installed Microsoft Word is the full-product build specialization inherited from M0009.

M0010 authoritative Tier-3 compatibility additionally requires installed LibreOffice.

Where portable components can remain portable without distorting the architecture, they should.

## Trust and safety

The Word renderer processes trusted YADG-authored DOCX and owns its automation instance.

Realistic compatibility fixtures are synthetic and redistribution-safe.

The publisher does not execute content and defines no remote credential protocol.
