# Changelog

## [Unreleased]

### YADG 1.1.0

YADG 1.1 focuses on authoring ergonomics, template discoverability, and explicit best-effort recovery without weakening strict production validation.

#### Added

- `yadg init` to bootstrap a new workspace without overwriting existing owned paths.
- `yadg check --list` to show discovered Markdown sources, DOCX templates, and referenceable semantic objects.
- `yadg inspect styles` to discover concrete Word style IDs, UI names, aliases, visibility and relevant numbering capability.
- `yadg inspect template` to explain placements, prototypes, semantic presentation roles, searchable Word locations, strict resolution, and YOLO fallback candidates.
- configurable Markdown policy for thematic breaks and inline code.
- character-style presentation for inline code.
- unordered/ordered list-item prototypes using real template paragraphs and numbering.
- generalized template-presentation resolution: concrete examples/prototypes, contextual template presentation, styles/numbering and compatibility defaults form one semantic model.
- human-searchable DOCX diagnostics with template context and nearby visible text where available.
- invocation-local `--yolo` for `check`, `build`, and `render`.
- visible degradation diagnostics for every successful YOLO recovery.
- deterministic related-template-resource fallback and conservative built-in real Word list numbering.
- visible YOLO placeholders for missing figures or failed Mermaid producers.
- visible preservation of unresolved value/reference identity instead of invented content.
- Word ↔ LibreOffice renderer fallback under explicit YOLO, with requested/actual renderer provenance.

#### Changed

- style selectors now resolve exact internal ID first, then exact Word display name, then alias.
- ordered-list presentation accepts real ordered Word numbering beyond decimal, including valid Roman/alphabetic formats.
- list/caption/prepared-table prototypes are treated as concrete presentation examples rather than isolated special cases.
- successful `build` replaces the complete current `YadgPreWords` generated result set.
- successful `render` replaces the complete current finalized result set; stale finalized DOCX/PDF artifacts do not survive.
- a successful Word render removes stale PDFs from an earlier LibreOffice render.
- publishing is downstream-only: it copies finalized `YadgWords/*.docx` without parsing current Markdown/templates or executing producers.
- recursive Markdown discovery prunes dot directories and YADG input/generated directories.
- linked child directories are not recursively followed, while ordinary accessible file links are not rejected merely for being links.
- diagnostics prefer workspace-relative paths and recognizable Word context.
- authoring/user documentation is expanded around a complete workspace example, templates/prototypes, inspection, and YOLO.

#### Compatibility and release notes

- existing strict 1.0 workspaces remain the compatibility baseline; YOLO is opt-in.
- workspace schema remains version 1.
- YADG remains a framework-dependent .NET tool.
- the 1.1 package targets `net10.0` and x64 while the authoritative full-product validation/support locus remains Windows x64.
- Microsoft Word rendering still requires installed/activated desktop Word in an interactive logged-on Windows session.
- YADG continues to use late-bound `Word.Application` and ships no Office interop wrapper binaries.
- LibreOffice remains the default renderer and continues to produce DOCX plus PDF.
- PDF publication is not part of 1.1; `publish` delivers finalized DOCX only.
- external Mermaid producers remain trusted user-configured executables and are not sandboxed by YADG.

The release date is intentionally not assigned here. Convert this prepared entry to `## 1.1.0 — <actual release date>` only during an explicitly authorized publication action.

## 1.0.0

Initial YADG 1.0 release:

- one `Yadg` .NET tool exposing `yadg check`, `build`, `render`, and `publish`;
- template-first Markdown/DOCX authoring with relative heading composition and document-wide visible value substitution;
- structured tables, prepared table rows, figures, captions, references, and trusted external Mermaid production;
- LibreOffice and interactive Microsoft Word finalization paths;
- explicit finalized-DOCX filesystem publishing.

YADG 1.0 is framework-dependent, Windows-only as the supported full-product locus, does not publish PDF, and does not provide unattended/server-side Word automation.
