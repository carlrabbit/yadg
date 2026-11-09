# Cross-format authoring consistency (M0019)

## Authority and ownership

This specification specializes existing YADG authoring policies for Word/DOCX and PowerPoint/PPTX. It is authoritative for the new M0019 behavior, and supersedes older format-specific statements only where explicitly identified. Existing `AUTHORING-RESILIENCE.md`, `CONTENT-PRODUCERS.md`, `WORKSPACE-VALUES.md`, `AUTHORING-EXPERIENCE.md`, `DOCUMENT-COMPOSITION.md`, `PRESENTATION-AUTHORING.md`, `TEMPLATE-PRESENTATION.md`, and `WORD-REFERENCES.md` remain authoritative for unaffected semantics.

Shared services own interpretation of common authoring policies, producer invocation, derived image asset lifecycle, value resolution, and degradation reporting. DOCX and PPTX retain separate Markdown structural models, semantic-ID scopes, Office-independent authoring backends and template-owned presentation. Do not replace the existing Word parser with the presentation parser, create a universal Office AST, or bring Office COM/UNO into generation.

This is a constrained extension of the existing external producer boundary, not a generic plugin framework.

## Shared Markdown policies

The existing schema-v1 `markdown.thematicBreak: error|ignore` and `markdown.codeInline: error|ignore|style` apply equally to configured DOCX and PPTX sources, including PPTX slide titles, text regions and table cells where the syntax can occur.

- `thematicBreak: error` rejects a thematic-break block; `ignore` omits the block without dropping adjacent content. Under `--yolo` an `error` thematic break may be omitted with one recorded degradation, as it is in DOCX.
- `codeInline: error` rejects code spans. `ignore` retains code characters as plain text; `style` retains them with a document/presentation template-owned code role, not a hardcoded generic monospace font.
- For PPTX `codeInline: style`, the selected text-region prototype containing the code requires an explicitly styled run whose *sole literal text* is `{{codeInline}}`, embedded in a formatting-example paragraph of that same region. That run is a role sample removed from generated slides, not a region marker or output text. The marker must be unique within that region. The run's actual character properties are the required exemplar; a missing/ambiguous sample fails strict validation only when that region's source contains styled code. The code-role sample can coexist with paragraph, heading and list samples. For title/table cells containing code, use their own native authored run exemplar in the corresponding title/cell prototype (same marker and uniqueness scope). Existing PowerPoint templates without code content need no change; templates with code content previously displayed plain code and require a code sample for true styled semantics. Document this deliberate compatibility tightening.
- Under explicit `--yolo`, `codeInline: error` recovers by preserving plain code text; a missing required style exemplar with `style` likewise recovers as plain text, each with one logical degradation per source occurrence. In strict mode no silent style fallback.
- Nested strong/italic/code and line breaks must preserve their intended combination and authored words. PPTX text regions support strikethrough using run properties supplied by the supported Office drawing model; semantics may compose with other inline emphasis. Failing to represent required formatting in strict mode is diagnostic, not silent content loss.
- External `http`, `https` and `mailto` links in PPTX text are native clickable hyperlinks retaining authored label text. Unsupported target kinds, including local Markdown fragments with no defined slide anchor, retain the authored label *unlinked* with an explicit non-degradation warning; no guessed navigation target. Existing Word link and generated-reference semantics remain unchanged.
- Blockquotes, literal fenced/indented code and grid tables remain DOCX-only in this milestone. PPTX typed regions still reject unsupported blocks in both strict and YOLO modes: do not turn arbitrary blocks into text or invent slide/table geometry.
- Preserve existing DOCX representation for pipe/grid tables, blockquotes, fenced code, text diagrams, references and headings. A common diagnostic classification does not imply all renderable blocks are format-neutral.

## Values across source and template boundaries

`values` remain case-sensitive single-line string scalars, literal/nonrecursive, with unchanged ID and malformed/missing-token diagnostic rules. Add **source-authored** `{{value:<id>}}` substitution for ordinary textual content in both formats. For DOCX, this is additive: existing template-visible story substitution, including Word headers/footers/notes/shapes, remains unchanged and uses the same values. In PPTX, substitution applies to slide titles, content headings, paragraphs, lists, table header/body cells, and ordinary visible static text on *selected* prototypes, including static titles/text-box content, but never to structural YADG control markers, relationship targets, drawing metadata, shape names, hyperlink URLs or the unselected prototype slides.

Resolution is performed before authoring output but after identifying which scope is an authored value and which is a structural marker. An exact `{{value:id}}` token may be split across consecutive runs within one text paragraph/cell; it must not span unrelated scopes. Replacement inherits the run/style properties of its first logical tag character and retains surrounding formatting. There is no recursive expansion of nested value-looking strings. Missing IDs fail strict mode; under YOLO preserve the visible token and emit one explicit degradation per unresolved occurrence. Malformed value-token syntax is fatal in both modes.

No new numeric/reference/caption semantics for PPTX. `[@id]` remains a DOCX semantic reference, not a PPTX slide jump. All configured/shared values work in direct mode only when explicitly supplied via `--config`.

## Producer and resolved image pipeline

The established `producers.mermaid` schema, external process invocation contract, Bun/npm neutrality, pinned test tools and PNG validation remain unchanged. Extract/reuse one producer invocation and resolved-image boundary for both outputs; the contract does not require an internal class name or a persistent cache.

A PPTX `image` region accepts exactly **one** local Markdown PNG/JPEG image OR **one** fenced Mermaid block with the existing grammar:

    ```mermaid {#architecture caption="Architecture"}
    flowchart LR
      A --> B
    ```

The Mermaid block must have a stable ID unique within its **presentation deck** across ordered source files. It does not create a Word global figure/section ID, slide selector, PPTX caption number or cross-slide reference. The optional caption is semantic producer metadata only: no automatic visible slide caption. The generated PNG is inserted as an ordinary native editable picture into the region, preserving the existing contain/letterbox geometry and alt text; no change to prototype region types.

`check` must execute the producer when needed to validate a real output, without committing a normal artifact. `build` must execute against current content and must not reuse stale producer output. Producer outputs are isolated temporary assets and cleaned on success/failure. A nonzero exit, timeout, absent/invalid PNG, missing configuration or executable is a strict error with actionable source/region and process diagnostics. A producer may execute from a trusted explicit config; no sandbox/network-isolation guarantee or tool installation by YADG is implied.

Under `--yolo`, a missing/unavailable Mermaid producer or failed invocation, and a missing in-bound image asset, yield one **visible editable placeholder shape** occupying the designated PPTX image region, labeled with stable source identity and failure reason as applicable. This is not a fabricated image or an empty slide. Emit one logical degradation per failed image/producer source occurrence. Lexical asset escape, malicious/invalid path syntax, ambiguous source identity and invalid/ambiguous prototype geometry remain fatal. DOCX uses its existing visible figure placeholder behavior; do not alter its numbering/reference semantics. A failed producer may not modify committed output prior to successful staging/commit.

## Strict/YOLO coverage

`--yolo` remains invocation-local and explicit. Enable it for configured presentation `check`/`build` and direct `document`/`presentation` `check`/`build`; it does not change `YADG.md`, change the set of source files, or cause automatic fallback to another presentation template.

Permitted PPTX recoveries are limited to: a prohibited thematic break omitted while retaining adjacent blocks; code-span text preserved plain under `error` or missing code example; unresolved valid value token retained visibly; missing in-bound local image replaced by visible placeholder; missing/unavailable/failing Mermaid producer replaced by visible placeholder. Log a single degradation per logical recovery with code, original problem, fallback, and source/slide/region when known; success summaries show consistent counts.

Remain fatal in YOLO: malformed root/direct config, duplicate IDs, invalid selectors/regions/layouts, missing/ambiguous/invalid selected PPTX prototypes or text/table structures, unsupported blocks in typed regions, table shape/column/structural-fit incompatibility, invalid/corrupt Office packages, unsafe/escaping paths, and no useful requested output. A deck is not silently skipped on failure. In mixed workspaces, retain the existing independence of DOCX and PPTX result-set commits; a previous DOCX set may have committed before a failing PPTX stage. Do not claim cross-format transactionality.

A configured presentation-only workspace with at least one valid selected deck is a useful `check --yolo`/`build --yolo` target without a DOCX template. Direct mode has one requested document/deck: failure to produce it is nonzero even in YOLO. A successful build commits only fully generated outputs, never partially-authored files.

## Direct authoring contract

Add `yadg document check` and `yadg document build`. Keep `yadg presentation check/build` and extend both command families with optional `--config <file.yaml>` and `--yolo`. Direct commands require exactly one explicit template, one or more repeatable ordered `--source` arguments, and on build exactly one `--output`. A check rejects `--output`. For example:

    yadg document build --template template.docx --source a.md --source b.md --output out/report.docx --config authoring.yaml --yolo
    yadg presentation build --template template.pptx --source slides.md --output out/slides.pptx --config authoring.yaml --yolo

A direct DOCX build produces **one authored, unfinalized DOCX** from the existing Word authoring pipeline; it does NOT invoke Word/LibreOffice finalization, write `YadgPreWords/`, produce PDF, render or publish. If a template contains placements, the ordered explicit sources are merged using the current Word global semantic identity/reference rules. Existing DOCX template validation, value substitutions, asset checks, producer semantics and YOLO recovery apply. Direct PPTX remains the existing one-deck path.

Without `--config`, direct commands use ordinary default authoring policies, empty values and no producer executable; no file is inferred or read. Explicit `--config` is a local UTF-8 YAML document with root `yadg.version: 1` and optional `markdown`, `values`, `producers`, `diagrams` and `references` mappings using the **same existing types and key/value validation** as workspace schema v1; reject all other roots including `presentations` and `publish`. No Markdown body, implicit YAML scalar coercion, anchors, aliases or custom tags. The file path is resolved against caller working directory; relative producer executable paths and working directory are based on the caller working directory, NOT the config-file directory. Environment/PATH lookup and explicit source/template path behavior remain as existing direct PPTX. No ambient `YADG.md` is read, even if malformed, and no other workspace source/template/producer is discovered. Relative image assets follow existing direct PPTX effective-root (caller working directory) and source-file resolution constraints; direct DOCX follows the same lexical-root rule. No implicit asset-root widening for source paths outside cwd.

For direct authoring, explicit `--config` plus producer-requiring source authorizes executing that producer during `check`/`build`. An unused configured producer must not be executed. No configured publication or arbitrary tool hooks are permitted. Direct output is one explicit file, written through a temporary sibling and atomically replaced only on success; prior output is unchanged on failure, inputs cannot be overwritten, and temp resources are cleaned. Direct `document` and `presentation` commands do not mutate other generated result sets.

## Selected configured presentation

Support a second invocation of `yadg presentation check/build` with `--id <declared-id>` and optional `--workspace <dir>` (default current directory). This is **configured selected-deck mode**: it reads `YADG.md`, validates global declaration schema/identity/output collisions, then validates/generates only the named deck. It does not parse DOCX source content, execute unrelated producers, author DOCX, or validate unrelated deck sources/templates. Missing ID and malformed global configuration are fatal. `--id` must be combined with neither `--template`, `--source`, `--output` nor `--config`; these belong to direct mode. `--yolo` is allowed in either mode.

Selected build transactionally replaces only `YadgPresentations/<selected-output>`; all other deck outputs remain byte-for-byte unchanged, including stale decks. Selected check does not mutate any output. Unselected declaration source/template failures are not blockers when the declaration table itself is valid. Existing unselected decks remain covered by ordinary workspace-wide `yadg check/build`, which continue to validate/build the full configured set and replace the entire PPTX result set on success.

Extend `yadg check --list` with deterministic workspace-relative presentation declarations: IDs, template, ordered source paths and output. The list is human-readable, non-machine-stable, and still displays successfully discovered entries on an otherwise invalid workspace. Do not conflate presentation-only sources with DOCX source inventory.

## Scope exclusions and compatibility

- No broad architecture rewrite, generic plugin API, extra producer kinds, SVG/PDF external producer product, global producer cache, remote producers, or runtime Office dependencies in authoring.
- No new PPTX layout families, arbitrary mixed-region content, custom shapes, SmartArt, slide reference/numbering semantics, or post-flight overflow measurement.
- Do not change `yadg render`, `publish` or DOCX renderer fallback/publication contracts.
- Do not change existing global workspace result-set semantics or make the multiple output types globally atomic.
- Existing schema-v1 `YADG.md` files remain valid. Optional direct config is a separate explicitly selected input and is never a workspace marker. Additional value/source-code and code-sample capabilities require only the explicitly documented source/template updates.
- End-to-end user docs must distinguish strict vs YOLO, direct vs workspace, selected vs whole-set builds, producer execution trust, and the one-file direct DOCX output being unfinalized.
