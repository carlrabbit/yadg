# Terminology

## Authoring

Office-independent transformation of workspace/Markdown semantics and prepared templates into `YadgPreWords/*.docx`.

## Renderer / finalizer

Concrete document application/runtime that consumes authored DOCX and establishes current field/index/layout-dependent document state.

## LibreOffice renderer

Existing renderer using isolated LibreOffice automation. Produces finalized DOCX and the existing PDF side output.

## Microsoft Word renderer

M0009 renderer using the desktop Microsoft Word COM automation model on Windows. Produces finalized DOCX.

## Finalized document

Top-level DOCX under `YadgWords/` that has successfully passed through a configured renderer.

## Publication

Explicit copy of finalized workspace DOCX to a caller/configured delivery filesystem destination.

Publication does not rebuild, rerender, or alter document contents.

## Publication destination

Effective filesystem path chosen by `--publish-path` or, when absent, `YADG.md publish.path`.

## Workspace publication default

Optional `YADG.md publish.path` used when `publish` receives no CLI destination.

It is not a fixed product output directory.

## Delivery artifact

In M0009, a finalized DOCX copied to the publication destination.

PDF is not an M0009 delivery artifact.

## Office Interop / Word automation

Managed interaction with Microsoft Word's COM-based object model.

M0009 does not prescribe early-bound PIA versus late-bound COM so long as real Word behavior and build/runtime constraints are satisfied.

## Interactive Word locus

Logged-on Windows user session with a normal user profile and activated desktop Word, used for supported M0009 automation/validation.

## Human artifact-quality review

Milestone-scoped human inspection of a real finalized document when automated structural checks cannot decide presentation fidelity.
