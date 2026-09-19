# Terminology

## Authoring

Office-independent transformation of Markdown and prepared Word templates into structurally complete DOCX artifacts.

## Renderer / finalizer

Separate real-rendering component that evaluates field/layout-dependent state such as `SEQ`, `REF`, TOC, lists, pagination, and later PDF output.

## Template front matter

Optional visible YAML stored in a removable control region at the start of a Word template.

It maps YADG semantic roles to template-local existing styles/prototypes. It does not define product semantics or formatting.

## Template control region

Removable main-body prefix containing YADG front matter and prototypes. It does not appear in authored output.

## Prototype

Template-owned Word-native structure identified by a template-local prototype ID and cloned/adapted by YADG.

M0004 caption prototypes carry literal labels, `SEQ` fields/switches, punctuation, style, and formatting.

## Semantic cross-reference

Markdown inline `[@stable-id]` resolving to a referenceable semantic object and authored as a Word `REF` field.

Surrounding prose supplies human-facing labels.

## Numbered target

A uniquely rendered Word target whose numeric value can be referenced: in M0004, a numbered figure/table caption or uniquely rendered numbered section heading.

## Internal bookmark

Word bookmark generated in an authored artifact so Word fields can target a numbered object.

Bookmark names are implementation mechanics, not semantic stable IDs.

## Field result

Displayed/cached value associated with a Word field. After Office-independent `build`, it is non-authoritative until evaluated by a renderer.

## Stable ID

Explicit case-sensitive workspace-wide semantic identity, separate from template prototype IDs and Word bookmark names.

## Structured object

Referenceable semantic block such as a table or figure.

## Natural anchor

Semantic source position of a structured object in Markdown.

## Placement override

Unique direct table/figure tag that relocates a structured object for one template.

## Derived artifact

Rebuildable generated output such as authored/finalized DOCX, PDF, or validation evidence.

## Validation target

Concrete system whose behavior establishes evidence, such as a real DOCX package or installed Word runtime.

## Validation locus

Where validation executes, such as ordinary local/CI .NET or Windows with Microsoft Word.


## Authored DOCX / PreWord

The Office-independent DOCX produced by `build` under `YadgPreWords/`. It contains semantic content and Word-native field structures but does not claim evaluated layout/field results.

## Finalized DOCX / Word artifact

The renderer-processed DOCX produced under `YadgWords/` with renderer-evaluated field/index state for the validated renderer.

## Renderer ID

A stable CLI selector for a concrete rendering engine implementation. M0005 defines `libreoffice`.

## LibreOffice renderer

The M0005 renderer implementation that loads authored DOCX files in an isolated LibreOffice process, refreshes fields/indexes, saves finalized DOCX, and exports PDF.

## Renderer session

One isolated loaded-document state in a concrete renderer from which the finalized DOCX and PDF are produced.

## Review request

A milestone-owned durable request under `.review/pending/` defining human evidence and acceptance criteria.

## Review record

A durable human decision under `.review/records/` for one milestone-owned review. Completed records are historical evidence, not perpetual approval of future repository state.
