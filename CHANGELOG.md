# Changelog

## [Unreleased]

### YADG 1.3.0 release candidate

The 1.3.0 candidate consolidates the unreleased 1.2.0 feature line with cross-format authoring, template-based PowerPoint presentations, and shared SVG box diagrams. No publication date or external release is claimed. Release validation and human release approval remain pending.

#### Added

- Built-in `diagram` fences for Unicode/ASCII character-grid diagrams, deterministic SVG, nested boxes, arrows and conservative occlusion handling; opt-in `diagrams.autoDetect`.
- Template-based PPTX authoring from Markdown with 16 standard layouts, typed text/image/table regions, and native PowerPoint/Impress prototype templates.
- Named presentation decks, selected-deck commands, direct `presentation check/build` commands and explicit direct `--config`.
- Cross-format Markdown policies, literal source-authored values, Mermaid producer reuse, strict/YOLO diagnostics and visible recovery semantics.
- Built-in box diagrams in PPTX as vector SVG pictures inside template-owned image regions, preserving the shared DOCX diagram parser; no native component-shape editing.
- User guides for Markdown, DOCX templates, presentations, box diagrams, and the complete schema-v1 `YADG.md` configuration reference.

#### Changed

- Markdown link/reference semantics and Markdig pipe/grid-table authoring were expanded for Word documents.
- Existing Word/Writer and PowerPoint/Impress origin-specific compatibility and installed-tool validation were strengthened.
- `README.md`, package description and user documentation now cover DOCX and PPTX.

#### Compatibility and release notes

- Workspace configuration remains `YADG.md`, schema version 1; `workspace.yaml` is not supported.
- Diagram auto-detection remains disabled by default. Mermaid remains an explicitly configured external PNG producer. User-file SVG figures are not enabled.
- DOCX publishing remains downstream-only; presentation outputs are authored PPTX, without a `render` or `publish` path.
- Package remains a framework-dependent `net10.0`, x64 `.NET` tool with Windows x64 as the authoritative full-product validation environment.
- The prior 1.2.0 section was an **unreleased candidate**, not a dated release entry; its capabilities are included here rather than represented as an externally published version.

## 1.1.0 — 2026-09-30

### YADG 1.1.0

YADG 1.1 focuses on authoring ergonomics, template discoverability, and explicit best-effort recovery without weakening strict production validation.

#### Added

- `yadg init` to bootstrap a new workspace without overwriting existing owned paths.
- `yadg check --list` to show discovered Markdown sources, DOCX templates, and referenceable semantic objects.
- `yadg inspect styles` and `yadg inspect template` for concrete template resource discovery.
- configurable Markdown policy for thematic breaks and inline code.
- character-style presentation for inline code.
- list-item prototypes using real template paragraphs and numbering.
- generalized template-presentation resolution and human-searchable DOCX diagnostics.
- invocation-local `--yolo` for `check`, `build`, and `render` with visible degradation reporting.
- deterministic presentation fallback, visible missing-figure/Mermaid placeholders, unresolved value/reference preservation, and renderer fallback.

#### Changed

- style selectors resolve concrete internal IDs/names/aliases instead of relying on gallery assumptions;
- ordered-list presentation accepts real non-decimal ordered numbering;
- successful build/render replace their complete generated result sets;
- publishing is downstream-only and does not rebuild/revalidate authoring input;
- recursive Markdown discovery prunes YADG/dot directories and does not follow linked child directories;
- user documentation was expanded around authoring, template inspection, and YOLO.

## 1.0.0

Initial YADG 1.0 release:

- one `Yadg` .NET tool exposing `yadg check`, `build`, `render`, and `publish`;
- template-first Markdown/DOCX authoring with relative heading composition and document-wide visible value substitution;
- structured tables, prepared table rows, figures, captions, references, and trusted external Mermaid production;
- LibreOffice and interactive Microsoft Word finalization paths;
- explicit finalized-DOCX filesystem publishing.
