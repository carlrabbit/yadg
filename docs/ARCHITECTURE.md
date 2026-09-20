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

## Word authoring

Owns prepared DOCX inspection/transformation through Open XML.

It creates structurally complete DOCX artifacts without requiring an Office renderer.

## External content producers

Remain separate-process source adapters.

M0008 Mermaid output becomes an ordinary semantic figure before Word authoring.

## Renderer boundary

Renderers consume `YadgPreWords/*.docx` and produce finalized workspace artifacts.

### LibreOffice

Existing isolated UNO renderer.

Produces finalized DOCX and its existing PDF side output.

### Microsoft Word

M0009 adds a Windows Microsoft Word automation specialization.

Word-specific COM types/code are isolated from authoring-core and Open XML authoring.

The concrete binding approach (PIA/interop metadata versus late-bound COM) is an implementation detail so long as the supported real-Word contract is satisfied.

The full Word-enabled product is allowed to become Windows-build-specific.

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

This separation allows publication to run in automation/CI even though Word rendering itself is validated only in an interactive Windows user session.

## Dependency direction

Semantic/core and Open XML authoring must not gain renderer-specific COM dependencies.

Renderer-specific runtime code depends inward on the common renderer contract.

Publishing depends only on workspace/finalized-artifact conventions and filesystem operations.

## Platform specialization

Windows + installed Microsoft Word is a concrete specialization for the `word` renderer.

No project requirement says the full solution must remain Linux-buildable after M0009.

Where portable components can remain portable without distorting the architecture, they should.

## Trust and safety

The Word renderer processes trusted YADG-authored DOCX and owns its automation instance.

The publisher does not execute content and defines no remote credential protocol.
