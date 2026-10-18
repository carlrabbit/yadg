# Changelog

## [Unreleased]

### YADG 1.2.0

YADG 1.2 adds built-in source-grid text diagrams and makes the authoring documentation boundary explicit.

#### Added

- built-in `diagram` fenced blocks that convert source-authored Unicode/ASCII character-grid geometry into passive SVG embedded in DOCX;
- supported orthogonal box drawing, arrows, nested boxes, connector crossings, and conservative overlap/occlusion reconstruction;
- opt-in `diagrams.autoDetect` for strong diagram geometry inside unlabeled or exact lowercase `text` fenced blocks;
- anonymous natural rendering for auto-detected diagrams without stable IDs;
- a dedicated Markdown authoring guide covering both built-in text diagrams and Mermaid;
- a dedicated DOCX template authoring guide covering placements, styles, prototypes, prepared tables, captions/references, values, and inspection.

#### Changed

- README/getting-started documentation distinguishes the two diagramming choices: built-in source-grid diagrams versus externally produced Mermaid diagrams;
- historical `docs/user/AUTHORING.md` and `docs/user/TEMPLATES.md` paths become compatibility landing pages pointing to the dedicated authoring guides;
- release/consumer validation covers built-in text diagrams and real Mermaid production from the installed package.

#### Compatibility and release notes

- workspace schema remains version 1;
- text-diagram auto-detection is off by default, so existing workspaces keep their previous fenced-block behavior unless they opt in;
- explicit `diagram` fences are built in and require no external diagram executable;
- Mermaid remains a separately configured trusted external producer and continues through its PNG figure path;
- source-file figures remain PNG/JPEG; YADG's internal SVG generation does not enable arbitrary source SVG figures;
- package/runtime platform policy remains the YADG 1.1 model: framework-dependent `net10.0`, x64, with Windows x64 as the authoritative full-product validation locus.

Do not assign a 1.2.0 release date or publish externally as part of M0015. Publication is a separate explicitly authorized operation.

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
