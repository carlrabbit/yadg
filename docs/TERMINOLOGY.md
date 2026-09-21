# Terminology

## Authoring

Office-independent transformation of workspace/Markdown semantics and prepared templates into `YadgPreWords/*.docx`.

## Template outline context

The nearest preceding prepared-template body paragraph with an effective Word outline level, used by M0010 to rebase headings for a `section`/`content` placement.

If none exists, the context level is zero.

## Source heading level

The Markdown heading level, 1 through 6, carried by the semantic Markdown model.

## Effective Word heading level

The template-relative Word outline/style role, 1 through 9, calculated during authoring for an emitted Markdown heading.

It is not stored as a Markdown semantic property.

## Heading rebasing

Mapping source Markdown heading hierarchy onto the template outline context while preserving source level deltas.

## Word story

A Word document text surface stored in a distinct document part/container, such as main body, header, footer, footnote, endnote, or comment.

The term is used for product explanation; users are not required to know OOXML story boundaries to use workspace values.

## Visible-text value substitution

Literal replacement of `{{value:<id>}}` in ordinary visible template text across the supported stories and text-box containers.

It does not create Word structures and does not edit field instructions or package metadata.

## Application-produced fixture

Committed DOCX compatibility template that originated as a new document and was populated/saved through Microsoft Word or LibreOffice Writer rather than synthesized by OpenXML SDK.

## Fixture provenance

Repository evidence recording the originating application/version/platform and SHA-256 of an application-produced fixture.

## Renderer / finalizer

Concrete document application/runtime that consumes authored DOCX and establishes current field/index/layout-dependent document state.

## LibreOffice renderer

Renderer using isolated LibreOffice automation. Produces finalized DOCX and the existing PDF side output.

## Microsoft Word renderer

Renderer using desktop Microsoft Word COM automation on Windows. Produces finalized DOCX.

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

A finalized DOCX copied to the publication destination.

PDF is not an M0009/M0010 delivery artifact.

## Office Interop / Word automation

Managed interaction with Microsoft Word's COM-based object model.

## Interactive Word locus

Logged-on Windows user session with a normal user profile and activated desktop Word, used for supported Word automation/validation.

## Human artifact-quality review

Milestone-scoped human inspection of real finalized documents when automated structural checks cannot decide presentation fidelity.
