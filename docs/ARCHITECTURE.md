# Architecture

## Objective

YADG separates workspace/content semantics, external source transformation, Word package transformation, template-owned presentation, and document rendering.

## `authoring-core`

Owns workspace discovery/front matter, Markdown parsing, OOXML-free semantic objects, stable IDs/references, and producer-independent semantic validation.

M0008 Mermaid blocks enter the semantic model as generated figures without leaking process/OOXML types into core.

## `content-producer` boundary

An external producer is a separate process invoked through a constrained adapter.

M0008 provides:

```text
Markdown Mermaid block
        |
        v
Mermaid producer adapter
        |
        +--> configured executable + argument prefix
        |
        +--> temporary .mmd input
        +<-- temporary .png output
        |
        v
semantic/generated figure asset
```

The process runner owns executable resolution, argument passing, bounded execution, stdout/stderr capture, temporary files, cleanup, and result diagnostics.

The Mermaid adapter owns Mermaid-specific `-i`/`-o` invocation and PNG validation.

The producer boundary does not receive Word/OOXML structures.

## `word-authoring`

Consumes the generated figure through the existing image/placement/caption/reference path.

No Mermaid-specific Word logic belongs here beyond the semantic figure being available with a valid PNG asset.

## Renderer

LibreOffice or any future renderer sees an ordinary embedded figure.

The renderer neither executes Mermaid nor knows how the figure was produced.

## Product versus test runtime

Bun is not a product dependency.

Product workspaces configure an executable and argument prefix.

Repository Tier-3 integration tests independently download the Bun version pinned in `eng/test-tools.json` and use Bun one-shot package execution of the pinned Mermaid CLI to establish a real external-producer target.

## Trust boundary

Workspace producer configuration is executable-code configuration.

YADG provides no sandbox. Process separation limits coupling, not authority of the configured executable.

## Dependency direction

Core semantic types remain independent of process runners and OOXML.

External process integration belongs in an authoring-side producer component that may be orchestrated by CLI/build/check.

Word authoring consumes validated generated assets; it does not initiate package installation.
