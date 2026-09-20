# Publishing Specification

## Status

Authoritative for the `publish` command introduced by M0009.

## Purpose

Publishing is the explicit boundary between finalized workspace artifacts and delivered documents.

M0009 publishes finalized DOCX only. It does not define DMS integration, release archives, manifests, signing, remote stores, or PDF delivery.

## Workspace artifact roles

```text
YadgTemplates/*.docx   template source
YadgPreWords/*.docx    authored intermediate DOCX
YadgWords/*.docx       finalized DOCX
YadgPdfs/*.pdf         existing LibreOffice renderer by-product
publish destination    delivered DOCX
```

Only top-level `YadgWords/*.docx` files are M0009 publish inputs.

## CLI

```text
yadg publish [--workspace <path>] [--publish-path <path>]
```

`--workspace` keeps its existing semantics.

Destination precedence:

```text
CLI --publish-path
    overrides
YADG.md publish.path
    otherwise
error
```

`publish` never invokes `build` or `render`.

## Workspace configuration

Root `YADG.md` front matter may contain:

```yaml
publish:
  path: ./Published
```

The `publish` mapping contains exactly `path`.

`path` is a non-empty string. Unknown `publish` keys are errors.

A workspace may omit `publish`; this remains valid for `check`, `build`, and `render`.

## Path resolution

Absolute publication paths are normalized and used directly.

Relative publication paths, from either CLI or configuration, resolve relative to the workspace root.

The destination may be inside or outside the workspace.

It must not resolve to, or underneath:

```text
YadgTemplates
YadgPreWords
YadgWords
YadgPdfs
```

A workspace-local path such as `./Published` is valid.

M0009 defines only ordinary filesystem destinations. UNC/mounted paths work only insofar as the operating system exposes them as filesystem paths.

## Inputs and outputs

At least one top-level finalized DOCX must exist in `YadgWords/`.

Publishing does not recurse.

Each input preserves its basename:

```text
YadgWords/<name>.docx
    ->
<publish-path>/<name>.docx
```

All top-level finalized DOCX files are published. There is no per-document filter in M0009.

The destination directory is created when necessary after preflight succeeds.

Existing same-name destination files may be replaced. Unrelated destination files are preserved.

M0009 does not clean stale destination files.

## Failure and commit behavior

Before changing delivered files, `publish` validates at least:

- workspace conventions;
- effective destination;
- destination is not a reserved YADG directory/descendant;
- finalized input directory;
- at least one publishable DOCX;
- source inputs are regular files;
- destination can be created/accessed sufficiently to begin publication.

Each output is copied through a temporary file in the destination filesystem and committed to the final filename only after the copy succeeds.

Rollback of already committed sibling files after a later sibling failure is not required.

Temporary/incomplete files must not be presented as successful delivery and are cleaned best-effort.

## Validation interaction

`check` validates optional `publish.path` schema/syntax but does not require the destination to exist or be writable.

`build` and `render` do not touch the publication destination.

Only `publish` performs destination runtime validation.

## Renderer independence

`publish` does not care whether `YadgWords` was produced by LibreOffice or Microsoft Word.

The publisher copies bytes and does not modify or recalculate document contents.

## PDF

`YadgPdfs` remains an existing LibreOffice output.

PDF is not a M0009 publish input and no V1 publication guarantee is made for PDF.

## Deferred behavior

M0009 does not define PDF publication, manifests/hashes, ZIP/release packaging, DMS integration, remote storage APIs, signing, filtering/renaming, stale-output cleanup, publication history, or implicit build/render.
