# Publishing Specification

## Status

Authoritative for the `publish` command after the M0012 correction.

## Purpose

Publishing is the explicit downstream boundary between finalized workspace artifacts and delivered documents.

It publishes finalized DOCX only.

It does not define DMS integration, release archives, manifests, signing, remote stores, or PDF delivery.

## Workspace artifact roles

```text
YadgTemplates/*.docx   template source
YadgPreWords/*.docx    authored intermediate DOCX
YadgWords/*.docx       finalized DOCX
YadgPdfs/*.pdf         renderer-specific generated PDF
publish destination    delivered DOCX
```

Only top-level `YadgWords/*.docx` files are publish inputs.

## CLI

```text
yadg publish [--workspace <path>] [--publish-path <path>]
```

Destination precedence:

```text
CLI --publish-path
    overrides
YADG.md publish.path
    otherwise
error
```

`publish` never invokes `build` or `render`.

It also never invokes the authoring pipeline.

## Downstream independence

Publishing does not need current Markdown/template authoring state to be valid.

It must not:

- parse Markdown content;
- inspect or validate DOCX templates;
- validate authoring figures/assets;
- execute Mermaid or any external content producer;
- require list/style/prototype resolution;
- require an authoring `check` to pass.

An already-finalized `YadgWords` artifact remains publishable even if the author has since edited/broken current source/template inputs.

### Explicit CLI destination

When `--publish-path` is supplied, publication uses:

- selected workspace root;
- finalized `YadgWords/*.docx`;
- CLI destination;
- publication-specific path/file preflight.

It does not need to parse `YADG.md` merely to validate unrelated authoring configuration.

### Configured destination

When `--publish-path` is absent, read `YADG.md` only as needed to obtain:

```yaml
publish:
  path: ...
```

If configuration cannot be parsed sufficiently to obtain a valid destination, publication fails.

Unrelated authoring validation is not part of publication.

## Path resolution

Absolute publication paths are normalized and used directly.

Relative publication paths resolve relative to workspace root.

The destination may be inside or outside the workspace.

It must not resolve to, or underneath:

```text
YadgTemplates
YadgPreWords
YadgWords
YadgPdfs
```

A workspace-local destination such as `./Published` is valid.

Ordinary filesystem links/junctions are handled by normal OS filesystem behavior; YADG does not establish a hostile-filesystem sandbox around publication.

## Inputs and outputs

At least one top-level finalized DOCX must exist in `YadgWords/`.

Publishing does not recurse.

Each input preserves basename:

```text
YadgWords/<name>.docx
    ->
<publish-path>/<name>.docx
```

All top-level finalized DOCX files are published.

The destination is created when necessary after preflight.

Existing same-name destination files may be replaced.

Unrelated destination files are preserved.

Publishing does not clean stale destination files.

## Failure and commit behavior

Before changing delivered files, validate at least:

- workspace root/destination context needed for publication;
- effective destination;
- destination is not a reserved YADG generated/input directory/descendant;
- finalized input directory;
- at least one publishable DOCX;
- source inputs are usable files;
- destination can be created/accessed sufficiently to begin publication.

Each output is copied through a temporary file in the destination filesystem and committed only after the copy succeeds.

Rollback of already committed sibling files after a later sibling failure is not required.

Temporary/incomplete files are cleaned best-effort and are not presented as successful delivery.

## Validation interaction

`check` may validate optional `publish.path` schema as part of authoring/workspace configuration, but `publish` does not rerun authoring validation.

`build` and `render` do not touch publication destination.

Only `publish` performs destination runtime validation.

## Renderer independence

`publish` does not care whether `YadgWords` was produced by LibreOffice or Microsoft Word, including YOLO renderer fallback.

The publisher copies bytes and does not modify/recalculate document contents.

## Producer non-execution acceptance

Validation must prove that publication does not execute external producers.

A test should configure a producer whose invocation would be observable, place an existing finalized DOCX in `YadgWords`, invoke `publish`, and prove successful delivery without producer invocation.

Also prove explicit `--publish-path` can publish finalized output while current Markdown/template authoring is invalid.

## PDF

`YadgPdfs` is not a publish input.

No V1/V1.1 PDF publication guarantee is introduced here.

## Deferred behavior

This specification does not define PDF publication, manifests/hashes, ZIP/release packaging, DMS integration, remote storage APIs, signing, filtering/renaming, destination synchronization, publication history, or implicit build/render.
