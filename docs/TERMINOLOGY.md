# Terminology

## External content producer

A separate executable process that converts explicitly declared source content into a YADG-supported product without receiving or mutating Word/OOXML structures.

## Producer kind

A YADG-defined source/product contract implemented through an external command. M0008 defines `mermaid`.

## Producer configuration

Workspace front-matter configuration specifying the executable and argument prefix used for a producer kind.

## Producer product

Validated output returned to YADG semantics. M0008's Mermaid product is a PNG-backed semantic figure.

## Generated figure

A semantic YADG figure whose image asset is produced during `check`/`build` rather than read from a source image file.

## Mermaid block

A Markdown fenced block whose info string is `mermaid {#id ...}` and whose body is Mermaid source.

## Producer invocation

One bounded external process execution using isolated temporary input/output files.

## Test-tool manifest

`eng/test-tools.json`, containing exact versions of external tools/packages used by authoritative integration tests.

## One-shot package execution

Running an npm package executable without adding it as a repository dependency/global install. M0008's authoritative integration target uses Bun `x`/`bunx` semantics.

## Semantic figure

A stable-ID figure participating in YADG natural/direct placement, captions, numbering, and references regardless of whether its image comes from a file or producer.

## Workspace value

Source-controlled scalar value from `YADG.md`, distinct from producer configuration and semantic object IDs.

## Authoring

Office-independent transformation of workspace metadata, Markdown semantics, producer products, and prepared DOCX templates into authored DOCX.

## Renderer / finalizer

Separate engine phase evaluating field/index/layout state after authoring.
