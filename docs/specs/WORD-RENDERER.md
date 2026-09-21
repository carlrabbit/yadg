# Microsoft Word Renderer Specification

## Status

Authoritative for the Microsoft Word renderer introduced by M0009.

## Purpose

Microsoft Word is the native application for the DOCX templates YADG targets.

M0009 adds a real Microsoft Word renderer/finalizer for field, index, pagination, and final document-state evaluation.

It consumes already-authored `YadgPreWords/*.docx`; it does not perform Markdown parsing or template authoring.

## Renderer ID and CLI

Renderer ID:

```text
word
```

Invocation:

```text
yadg render --renderer word [--workspace <path>]
```

The existing default remains `libreoffice`.

`--renderer-path` remains LibreOffice-specific and is invalid with `--renderer word`.

## Runtime locus

The authoritative Microsoft Word runtime is:

- Windows;
- desktop Microsoft Word installed and COM-registered;
- an interactive logged-on user session with a normal user profile;
- Office activation/first-run prompts already resolved.

YADG does not claim support for Word automation from Windows services, SYSTEM tasks, ASP/ASP.NET, DCOM server contexts, or other non-interactive/server-side execution.

The YADG Word renderer processes documents serially.

## Build and binding locus

M0009 no longer requires the complete YADG solution to remain Linux-buildable.

The authoritative build/validation locus for the Word-enabled product is Windows.

Word automation must remain isolated from authoring-core and Open XML authoring.

The corrected M0011 V1 binding is built-in late-bound COM using the version-independent `Word.Application` ProgID and `Activator.CreateInstance`. The renderer has no Office `COMReference`, generated interop assembly, Office interop package, handwritten interface, or source-generated wrapper dependency.

Microsoft Word is a runtime prerequisite, not a build/package prerequisite.

Preserving non-Windows full-solution compilation is welcome but is not an acceptance criterion.

## Application ownership

Each Word render invocation creates and owns its Microsoft Word application automation instance.

It must not attach to an unrelated already-running interactive Word instance.

Normal automation is non-visible and suppresses ordinary alerts where Word permits.

Every opened document must be closed and the owned Word application quit on success and best-effort on failure.

The implementation must avoid leaving its owned `WINWORD.EXE` behind after normal handled success/failure.

## Inputs and outputs

The Word renderer processes top-level:

```text
YadgPreWords/*.docx
```

and writes corresponding:

```text
YadgWords/*.docx
```

It never modifies `YadgPreWords`.

The M0009 Word renderer does not create PDF and does not modify `YadgPdfs`.

LibreOffice retains its existing DOCX/PDF behavior.

## Finalization semantics

For every document, Word opens a working copy through its object model and establishes current displayed values for YADG-supported structures.

Observable requirements:

- `SEQ` caption numbering is current;
- semantic `REF` values are current;
- numbered section references are current;
- template-owned TOC is refreshed when present;
- template-owned list-of-figures/list-of-tables structures are refreshed when present;
- other fields relevant to those structures are updated as needed;
- pagination is recalculated sufficiently for current field/index results;
- the finalized DOCX is saved by Microsoft Word.

The exact order of Word object-model calls is implementation-owned. Correctness is defined by the resulting artifact.

The renderer must not alter YADG semantic contracts to accommodate Word.

## Save and commit behavior

Finalization occurs on an isolated working copy, not the authored input.

A finalized DOCX is committed to `YadgWords` only after Word successfully opens, refreshes, and saves it.

Existing corresponding finalized DOCX may be replaced after preflight.

Unrelated files in `YadgWords` are preserved.

Temporary/failed files must not be presented as successful output.

Cross-document rollback after a later sibling failure is not required.

## Runtime detection and provenance

Before normal output modification, the renderer validates that Word can be instantiated and interrogated.

Real integration/review evidence records at least:

- Windows/OS version;
- Word/Office version observable through automation;
- renderer ID;
- source authored artifact identity/hash where practical;
- finalized artifact identity/hash.

Activation alone is not proof of support; the required real-runtime scenarios must pass.

## Security, prompts, and bounded execution

M0009 processes trusted `.docx` authored by YADG.

Automation should suppress macro execution and interactive alerts through available Word/Office controls where practical.

Macro-enabled input is not part of the supported contract.

A modal prompt preventing bounded completion is a renderer failure.

Use bounded operation/watchdog behavior sufficient to avoid an indefinitely hanging CLI. Exact timeout mechanics/values are implementation-owned.

## Failure diagnostics

Actionable failures include:

- Word COM activation unavailable;
- document open failure;
- blocking automation/timeout;
- refresh/finalization failure;
- save failure;
- inability to close/quit the owned automation cleanly.

Diagnostics identify document/stage and retain useful COM/HRESULT context without exposing document contents unnecessarily.

## Relationship to LibreOffice

`libreoffice` remains supported and remains the default renderer in M0009.

`word` is an additional renderer and the V1 fidelity target.

No byte/layout equivalence is claimed between Word and LibreOffice.

## Deferred behavior

M0009 does not define Word PDF export, unattended/server-side Word support, remote Word automation, concurrent Word automation, automatic Office installation/licensing, macro-enabled inputs, Word Online/Graph rendering, or renderer equivalence guarantees.
