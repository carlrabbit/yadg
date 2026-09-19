# Rendering Specification

## Status

Authoritative for renderer/finalizer behavior introduced by M0005.

## Renderer boundary

Authoring and rendering are separate phases.

```text
YadgTemplates/ + Markdown
        |
        | yadg build
        v
YadgPreWords/*.docx
        |
        | yadg render
        v
YadgWords/*.docx
YadgPdfs/*.pdf
```

`build` remains Office-independent.

`render` consumes authored DOCX artifacts; it does not parse Markdown, resolve template tags, or rebuild authoring semantics.

## Initial renderer

The first supported renderer ID is:

```text
libreoffice
```

LibreOffice is an implementation of the renderer boundary, not a substitute name for Microsoft Word. YADG does not claim that LibreOffice and Microsoft Word have identical layout or field behavior.

A later renderer may implement the same artifact contract with Microsoft Word or another engine.

## CLI

M0005 defines:

```text
yadg render [--workspace <path>] [--renderer libreoffice] [--renderer-path <path>]
```

`--workspace` defaults to the current directory.

`--renderer` defaults to `libreoffice`. Other renderer IDs are unsupported in M0005.

`--renderer-path` optionally identifies the LibreOffice command executable. An explicit path wins. Without it, YADG searches the process `PATH` for an appropriate LibreOffice command.

Runtime installation location is not workspace or template presentation configuration.

## Inputs and outputs

`render` processes top-level `.docx` files directly inside `YadgPreWords/`; it does not recurse.

At least one authored DOCX is required.

For every successful input:

```text
YadgPreWords/<name>.docx
  -> YadgWords/<name>.docx
  -> YadgPdfs/<name>.pdf
```

`render` never modifies `YadgPreWords` inputs.

Existing corresponding finalized DOCX/PDF outputs may be replaced after preflight validation succeeds. Unrelated files in output directories are not deleted.

`render` does not implicitly invoke `build`.

## Finalization semantics

For each input, YADG loads the authored DOCX in one isolated LibreOffice document session and, before final outputs are committed:

1. refreshes/recalculates imported text fields required by M0004, including caption sequence and semantic-reference fields;
2. updates document indexes exposed by Writer, including existing template-owned TOC/list structures;
3. performs any additional refresh required for dependent field/index values to become current;
4. saves the refreshed document as DOCX;
5. exports PDF from the same refreshed loaded document state.

The finalized DOCX and PDF therefore represent the same renderer session/state.

Plain `--convert-to docx` round-tripping of an authored DOCX is not sufficient evidence of M0005 finalization. Correctness is defined by observable refreshed results and preserved structures.

## M0004 field compatibility

The LibreOffice renderer must establish real-runtime compatibility with the Word-native structures authored by M0004:

- `SEQ` caption fields;
- bookmarked figure/table reference targets;
- `REF` numeric references;
- paragraph-number section references;
- existing template-owned document indexes such as TOC and figure/table lists when present.

A renderer failure to correctly import, refresh, preserve, save, or export these structures is a product failure for the LibreOffice renderer; it must not be hidden by rewriting M0004 semantic contracts during implementation.

## Field values after render

Unlike `build`, a successful `render` claims current renderer-evaluated field/index results for the LibreOffice output artifact.

The finalized DOCX must retain usable field/index structures where LibreOffice's DOCX export supports them and must carry current displayed results for the validated M0004 scenarios.

The PDF must display the corresponding evaluated values.

## LibreOffice process isolation

Every `render` invocation uses a dedicated temporary LibreOffice user profile rather than the user's ordinary LibreOffice profile.

The renderer launches LibreOffice headlessly with recovery/UI disabled as appropriate for automation and uses a local-only UNO/API connection.

The renderer must not attach to or depend on an already running interactive LibreOffice instance.

Temporary profile/process resources are cleaned up on success and best-effort on failure.

The exact UNO connection transport, startup handshake, and implementation library are implementation mechanics provided they are local-only and satisfy the process-isolation contract.

## Runtime capability and version evidence

M0005 uses capability-based runtime detection rather than a hard-coded minimum LibreOffice version.

The renderer must obtain and report the concrete LibreOffice version used for a render invocation.

Milestone integration evidence and human-review provenance must record:

- LibreOffice version;
- operating system/platform;
- renderer executable used;
- source authored artifact identity/hash where practical.

A version is not considered supported merely because it launches; the required M0004 field/index and DOCX/PDF behaviors must pass real-runtime validation.

## Failure semantics

Before normal outputs are modified, `render` validates:

- workspace/input conventions;
- renderer ID;
- renderer executable availability;
- ability to obtain runtime version;
- basic process/API startup capability.

Runtime/open/refresh/save/export failures produce actionable diagnostics and non-zero exit status.

If an external renderer failure occurs after output writing begins, transactional rollback of already completed sibling outputs is not required, but temporary/incomplete output files must not be presented as successful finalized artifacts.

The implementation must use bounded startup/operation waits rather than hanging indefinitely. Exact timeout values are implementation-owned unless future evidence requires product-level configuration.

## PDF export

M0005 uses LibreOffice Writer PDF export semantics. Default PDF export behavior is sufficient; M0005 introduces no workspace-configurable PDF filter options.

PDF encryption, signing, watermarks, PDF/A policy, page ranges, and other publication-specific PDF settings are deferred.

## Security and external content

The renderer processes `.docx` inputs only.

M0005 does not introduce macro-enabled document support, remote-document fetching, or renderer network dependencies.

The isolated runtime must not require user credentials.

## Human-visible correctness

Automated structural and runtime checks cannot fully decide whether template presentation survives rendering acceptably.

M0005 therefore has a blocking milestone-scoped human artifact-quality review defined by its milestone and `.review/pending/HR-M0005-01.md`.

## Deferred behavior

M0005 does not define:

- Microsoft Word renderer support;
- semantic equivalence between renderers;
- `publish`;
- DMS integration;
- packaging/distribution;
- configurable PDF export policy;
- automatic installation/update of LibreOffice;
- remote renderer services.
