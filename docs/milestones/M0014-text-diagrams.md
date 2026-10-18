# Milestone — M0014 Text Diagrams

## Execution Profile

| Field | Value |
|---|---|
| Lifecycle state | ready |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna |
| Baseline executor readiness | confirmed |
| Decision preservation | confirmed |
| Execution tractability | confirmed |
| Scope size | medium |
| Implementation autonomy | high |
| Documentation sync | deferred except direct user-facing feature docs |
| Focused validation | fenced syntax, grid parser, primitives, overlap/containment, auto-detection, semantic figure integration, SVG authoring |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real Word + LibreOffice consumption of generated SVG-backed diagrams |
| Validation locus/platform | Windows; installed Word and LibreOffice for Tier 3 |
| Consumer/release validation | not applicable |
| Human review | blocking artifact-quality review `HR-M0014-01` |

## Goal

Add a small built-in diagram generator for readable text/ASCII/Unicode diagrams while preserving the source as understandable Markdown.

The author supplies layout directly as a character grid. YADG interprets renderable drawing primitives and text, reconstructs conservative box containment/occlusion where unambiguous, and emits a vector diagram without graph-layout semantics.

The milestone also adds an opt-in convenience mode that auto-detects diagrams only inside unlabeled or exact `text` fenced blocks by asking whether the normal diagram parser can extract sufficiently strong renderable geometry.

## Target State

- explicit fenced `diagram` blocks render as vector diagrams;
- explicit diagrams are identified semantic figures with existing placement/caption/reference behavior;
- diagram layout is authored by the source grid; YADG performs no graph layout, node layout, or edge routing;
- supported box drawing, orthogonal connectors, arrows, text, containment, and unambiguous occlusion/layering render correctly;
- generated diagram product is ephemeral SVG and is embedded into DOCX through an SVG-capable figure path;
- source-authored file figures remain PNG/JPEG only; M0014 does not generally enable arbitrary workspace SVG files;
- `YADG.md` schema v1 gains opt-in `diagrams.autoDetect: true|false`, default `false`;
- when enabled, only unlabeled fences and exact lowercase `text` fences are auto-detection candidates;
- auto-detection reuses the real diagram parser and requires strong renderable geometry; it is not a separate punctuation score;
- unlabeled auto-detected diagrams may render anonymously at their natural source position;
- an auto-detected `text` block with a stable ID uses ordinary semantic figure identity;
- false/weak candidates retain the pre-existing unsupported-fenced-block behavior rather than being guessed into diagrams;
- real Word and LibreOffice consume the resulting DOCX without repair and show representative diagrams plausibly.

## Scope

M0014 covers:

- built-in character-grid diagram parsing;
- Unicode box-drawing primitives;
- supported ASCII line/junction/arrow compatibility;
- text spans positioned on the grid;
- orthogonal line connectivity;
- arrows and junctions;
- closed rectangle detection;
- nested rectangle containment;
- conservative reconstruction of rectangle edges hidden by a higher visible rectangle;
- z-order inference from unambiguous occlusion;
- connectors crossing container boundaries without clipping;
- deterministic SVG generation;
- explicit diagram fenced-block syntax;
- optional auto-detection for unlabeled/`text` fences;
- workspace configuration for auto-detection;
- semantic figure integration for identified diagrams;
- anonymous natural rendering for unlabeled auto-detected diagrams;
- check/build validation and diagnostics;
- direct user documentation;
- real Word/LibreOffice validation and blocking visual review.

## Non-goals

M0014 does not implement:

- graph layout algorithms;
- automatic node placement;
- automatic edge routing;
- Mermaid/PlantUML compatibility;
- free-form vector drawing;
- diagonal lines;
- arbitrary curves;
- arbitrary shape types beyond the supported orthogonal primitives/rectangles;
- arbitrary z-index syntax;
- colors or per-shape styling syntax;
- a general SVG source-image feature;
- source-file SVG figures;
- external diagram producer execution;
- diagram editing inside Word;
- OCR or interpretation of raster images;
- auto-detection in normal Markdown paragraphs, block quotes, lists, tables, indented code, or language-tagged fences other than exact lowercase `text`;
- auto-detection when `diagrams.autoDetect` is absent or false;
- heuristic conversion of weak/ambiguous punctuation art;
- persistence or caching of generated SVG files.

## Decisions and Constraints

### Built-in diagram boundary

Text diagrams are built-in semantic authoring content, not an external content producer.

The pipeline is:

```text
Markdown fenced source
-> character grid
-> renderable primitive model
-> optional rectangle/occlusion reconstruction
-> SVG product
-> ordinary Word authoring / figure placement
```

No external process is required.

The parser and primitive model remain OOXML-free.

### Source controls geometry

The source grid is the layout authority.

YADG does not infer graph semantics such as "service", "node", "edge", "subsystem", or "swimlane".

Whitespace is geometrically significant.

Ordinary text remains ordinary text positioned on the character grid.

### Explicit syntax

Supported explicit form:

````markdown
```diagram {#system-flow caption="System flow"}
┌──────────┐
│ A        │
└────┬─────┘
     ▼
┌──────────┐
│ B        │
└──────────┘
```
````

Rules:

- language token is exactly lowercase `diagram`;
- stable ID is mandatory;
- ID uses the existing semantic object lexical form;
- optional `caption` is plain text;
- unknown attributes are invalid;
- empty source is invalid;
- explicit `diagram` always invokes diagram parsing and invalid/unrenderable geometry is a check error.

An explicit identified diagram is semantically a figure and uses existing natural/direct placement, caption, numbering, bookmark, and reference rules.

### Auto-detection configuration

Schema v1 gains optional:

```yaml
diagrams:
  autoDetect: true
```

`diagrams` accepts exactly `autoDetect`.

`autoDetect` is a YAML boolean.

Default is `false`.

Unknown `diagrams` keys or non-boolean values are configuration errors.

Existing schema-v1 workspaces remain unchanged when this mapping is absent.

### Auto-detection candidate fences

When `autoDetect` is true, candidates are limited to:

1. fenced blocks with an empty language/info string;
2. exact lowercase `text` fences, optionally carrying the same `{#id caption="..."}` attribute form.

No other language token is inspected for automatic diagram conversion.

`diagram` remains explicit and does not depend on the toggle.

A non-diagram candidate retains the repository's ordinary unsupported-fenced-block behavior. Auto-detection is not a generic code-block renderer.

### Detection by parser success and strong primitives

Auto-detection must use the same parser/primitive extraction path used by explicit diagrams.

It may convert a candidate only when the parsed result contains strong renderable geometry, such as at least one of:

- a complete or conservatively reconstructable rectangle;
- a connected line path terminating in a supported arrowhead;
- a branching/joining connected component using line/junction primitives;
- a multi-row connected Unicode box-drawing component above a small implementation-owned minimum size.

Isolated punctuation, a single `A -> B`, key/value records, ordinary prose alignment, or disconnected decorative characters are insufficient.

The exact numeric thresholds are implementation-owned and must be covered by regression tests. The contract is deliberately biased toward false negatives over false positives.

### Primitive vocabulary

M0014 must support at least these Unicode structural primitives:

```text
─ │ ┌ ┐ └ ┘ ├ ┤ ┬ ┴ ┼
→ ← ↑ ↓ ► ◄ ▲ ▼
```

ASCII compatibility must support connected uses of:

```text
- | + > < ^ v
```

ASCII punctuation is structural only when connectivity/geometry makes that interpretation valid; otherwise it remains text.

Implementation may normalize supported source glyphs into one connectivity representation.

### Geometry and connectivity

Each structural cell contributes directional connectivity to neighboring cells.

Text is grouped into spans where useful but retains source grid position.

Lines and boxes use fixed cell geometry; the visual renderer must not depend on graph-layout algorithms.

Connectors may cross a containing box boundary. Boxes are visual containment, not clipping regions.

### Rectangles, containment, overlap, and layering

Complete axis-aligned rectangles are recognized from connected borders.

A rectangle fully inside another rectangle establishes visual containment but does not imply domain semantics.

Partial overlap is supported when z-order can be inferred conservatively from occlusion.

A missing rectangle edge segment may be reconstructed only when:

- visible fragments establish the same aligned rectangle edge;
- the missing interval is geometrically covered by another recognized rectangle;
- the covering rectangle's visible border/body explains the interruption;
- reconstruction yields one unambiguous rectangle/z-order interpretation.

The covering rectangle is above the reconstructed rectangle.

Do not invent hidden geometry across uncovered gaps.

Ambiguous competing reconstructions are invalid for explicit diagrams and must not qualify for auto-detection.

Nested/overlapping boxes may be rendered in back-to-front z-order derived from containment/occlusion.

### SVG product

Generated diagrams use SVG as the ephemeral authoring product.

The SVG must:

- be deterministic for the same semantic input/configuration;
- have explicit finite dimensions/viewBox;
- use vector strokes for lines/boxes/arrows;
- place text consistently on the source grid;
- escape text/XML safely;
- preserve transparent background unless implementation requires a compatible explicit background;
- contain no scripts, external references, remote resources, embedded active content, or runtime network dependency.

M0014 may extend Word authoring to embed internally generated SVG.

This does not authorize arbitrary source-file SVG figures.

Generated SVG remains ephemeral and is not written as workspace source/cache.

### Anonymous auto-detected diagrams

A truly unlabeled auto-detected block has no stable source identity.

It renders only at its natural source position and:

- is not registered in the global semantic ID namespace;
- cannot be targeted by `{{figure:id}}`;
- cannot be referenced by `[@id]`;
- has no caption/numbering/bookmark;
- is not listed as a referenceable object by `check --list`.

This is a deliberate convenience surface, not generated identity.

A `text` candidate carrying `{#id ...}` that is successfully auto-detected becomes an ordinary identified semantic figure and follows the explicit figure rules.

No synthesized path/line-number ID may become public semantic identity.

### Existing Mermaid and file figures

Mermaid remains an external content producer and continues producing PNG under its existing contract.

PNG/JPEG file-backed figures remain unchanged.

M0014 must not route Mermaid through the text-diagram parser or alter Mermaid producer configuration.

### Strict/YOLO behavior

Strict mode follows the rules above.

M0014 does not add new YOLO recovery semantics for malformed explicit diagrams or ambiguous geometry. Existing generic visible-degradation behavior may apply only where already authorized by `AUTHORING-RESILIENCE.md`; implementation must not invent a new silent fallback.

## Baseline Executor Readiness

Planning has settled:

- built-in versus external-producer boundary;
- source-as-layout semantics;
- explicit fence syntax;
- ID/caption semantics;
- workspace configuration/default;
- exact auto-detection candidate scope;
- parser-based detection principle;
- conservative false-negative bias;
- minimum required primitive vocabulary;
- containment/overlap/occlusion semantics;
- no clipping at box boundaries;
- anonymous auto-detected behavior;
- SVG product boundary;
- relationship to existing figure/Mermaid behavior;
- validation and human-review policy.

Implementation owns concrete parser types/algorithms, numeric auto-detection thresholds, SVG element organization, cell metrics/font mechanics, diagnostic codes/messages, OpenXML mechanics for SVG embedding, test-file organization, and work-package decomposition, provided they satisfy this contract.

No material product decision is intentionally deferred.

## Execution Tractability

Create and maintain:

```text
.execution/M0014-text-diagrams.md
```

The planning-seeded ledger already contains every obligation and validation gate. Implementation decomposes the work into bounded work packages and maps them back to those rows.

## Required Authority

Read before implementation:

- `docs/specs/TEXT-DIAGRAMS.md`
- `docs/specs/STRUCTURED-CONTENT.md`
- `docs/specs/WORD-REFERENCES.md`
- `docs/specs/WORKSPACE-VALUES.md`
- `docs/specs/AUTHORING-EXPERIENCE.md`
- `docs/specs/CONTENT-PRODUCERS.md`
- `docs/specs/AUTHORING-RESILIENCE.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`

Do not use the external guide repository or planning conversation as implementation authority.

## Acceptance Criteria

| ID | Acceptance obligation |
|---|---|
| AC-01 | Exact lowercase `diagram` fences with mandatory stable ID and optional plain caption parse as identified semantic figures; invalid/empty/unknown-attribute forms fail. |
| AC-02 | Unicode box-drawing primitives, required arrow glyphs, connected ASCII compatibility primitives, and ordinary text are represented by one OOXML-free character-grid/primitive model. |
| AC-03 | Complete rectangles and nested containment render according to source geometry without graph-layout or domain-semantic inference. |
| AC-04 | Unambiguous occluded rectangle edges are reconstructed and rendered with the covering rectangle above the covered rectangle; uncovered or ambiguous gaps are not invented. |
| AC-05 | Orthogonal connectors, junctions, arrows, branches/joins, and connectors crossing box boundaries render without unintended clipping. |
| AC-06 | Generated diagrams produce deterministic, finite, passive/self-contained SVG with no script/external-resource/network behavior and no persistent generated-asset cache. |
| AC-07 | Identified diagrams use existing natural/direct figure placement and caption/reference semantics, including duplicate-ID rules and direct-placement suppression. |
| AC-08 | Internally generated SVG is authored into DOCX through a supported SVG path without enabling arbitrary source-file SVG figures. |
| AC-09 | `YADG.md` schema v1 accepts optional `diagrams.autoDetect` boolean, defaults it to false, rejects unknown/invalid diagram configuration, and preserves existing workspace behavior when absent. |
| AC-10 | With auto-detection disabled, unlabeled and `text` fences retain their pre-M0014 unsupported behavior; explicit `diagram` remains supported independently. |
| AC-11 | With auto-detection enabled, only empty-info and exact lowercase `text` fences are candidates; other language-tagged fences are never auto-detected. |
| AC-12 | Auto-detection reuses normal diagram primitive extraction and converts only strong connected renderable geometry, with regression coverage proving representative weak/non-diagram blocks are not converted. |
| AC-13 | An unlabeled auto-detected diagram renders only at its natural source position and remains anonymous/non-referenceable/non-listable; no public synthesized ID is created. |
| AC-14 | A successfully auto-detected `text` fence carrying a stable ID/caption becomes an ordinary identified semantic figure with normal placement/reference behavior. |
| AC-15 | Existing PNG/JPEG file figures and Mermaid external-producer behavior remain unchanged. |
| AC-16 | `check` and `build` validate diagram parsing/renderability before normal output mutation and surface actionable source-position diagnostics for explicit diagram failures/ambiguity. |
| AC-17 | Representative generated DOCX opens/finalizes successfully through both real Microsoft Word and real LibreOffice and preserves the generated SVG diagram content without repair/fatal conversion error. |
| DOC-01 | README/direct authoring documentation describes explicit diagrams, supported primitive scope, layering/occlusion limits, anonymous diagrams, auto-detection opt-in and its conservative behavior. |
| REV-01 | Human artifact-quality review `HR-M0014-01` approves representative explicit and auto-detected diagrams including nested and overlapping boxes in Word and LibreOffice outputs. |

## Evidence Cases

| ID | Parent obligation | Required evidence case | Why separate |
|---|---|---|---|
| EC-04a | AC-04 | one lower rectangle partially hidden by one higher rectangle reconstructs successfully | positive occlusion path |
| EC-04b | AC-04 | ambiguous/uncovered missing-edge geometry is rejected/not auto-detected | materially different failure path |
| EC-07a | AC-07 | natural identified diagram placement | natural-anchor path |
| EC-07b | AC-07 | direct `{{figure:id}}` placement suppresses natural emission | separate placement dispatch |
| EC-12a | AC-12 | Unicode box/arrow candidate auto-detects | high-confidence Unicode path |
| EC-12b | AC-12 | connected ASCII candidate auto-detects | separate ASCII interpretation path |
| EC-12c | AC-12 | weak text/key-value/`A -> B` candidates do not auto-detect | false-positive prevention path |
| EC-17a | AC-17 | Word finalizes/opens representative generated diagram DOCX | renderer-specific compatibility |
| EC-17b | AC-17 | LibreOffice finalizes/opens representative generated diagram DOCX | renderer-specific compatibility |

## Validation

| ID | Depth | Target | Locus/platform | Command/check | Proves | Expected evidence |
|---|---|---|---|---|---|---|
| VAL-01 | Tier 1 | diagram parser/geometry/SVG | ordinary repository test locus | focused tests chosen by implementation | AC-01, AC-02, AC-03, EC-04a, EC-04b, AC-05, AC-06 | deterministic parser/primitive/SVG assertions including overlap and invalid ambiguity |
| VAL-02 | Tier 1 | auto-detection/configuration | ordinary repository test locus | focused tests chosen by implementation | AC-09, AC-10, AC-11, EC-12a, EC-12b, EC-12c, AC-13, AC-14 | positive/negative candidate matrix plus schema/default behavior |
| VAL-03 | Tier 1/2 | semantic figure + DOCX authoring | Windows-capable repository tests | focused DOCX structural tests | EC-07a, EC-07b, AC-08, AC-15, AC-16 | SVG image relationship/drawing structure, placement/reference evidence, compatibility regressions |
| VAL-04 | Tier 2 | full repository | authoritative Windows repository locus | `./eng/validate.ps1` | AC-01..AC-16, DOC-01 as applicable to repo checks | complete portable repository validation pass |
| VAL-05 | Tier 3 | real renderer compatibility | Windows interactive user session with activated Word + installed LibreOffice | dedicated M0014 Tier-3 target created by implementation | EC-17a, EC-17b | runtime/version/result evidence and representative finalized artifacts |
| VAL-06 | Tier 5 | visual artifact quality | human review of VAL-05 outputs | `./eng/review-check.ps1 --milestone M0014` | REV-01 | signed/recorded approval for `HR-M0014-01` |

Tier 3 must exercise at least:

```text
explicit identified diagram
nested box
partially overlapping/occluded boxes
orthogonal arrow/connector
auto-detected unlabeled diagram
auto-detected identified text diagram
ordinary non-diagram text fence negative case in check validation
```

The representative explicit diagram should contain enough geometry to make line joins, text alignment, containment, z-order, and arrowheads visually inspectable.

Tier 3 may reuse the repository's existing real Word/LibreOffice capability and isolation policy. It must not provision or add new CI infrastructure merely to satisfy this milestone. If the required installed runtime is unavailable on the implementation host and no already-authorized equivalent locus exists, preserve all agent-resolvable work/evidence and terminate `BLOCKED` or `AWAITING HUMAN REVIEW` as appropriate rather than weakening the target.

## Human Review

Applicability: required.

Review class: `artifact-quality`

Canonical review ID:

```text
HR-M0014-01
```

Review subject:

- one representative M0014 diagram artifact finalized/opened with Microsoft Word;
- one representative M0014 diagram artifact finalized/opened with LibreOffice;
- both derived from the Tier-3 scenario and containing nested/overlapping boxes plus arrows/text.

Required evidence identifies the repository revision, source fixture/workspace, renderer/runtime version, and reviewed finalized artifacts.

Reviewer checks:

- boxes/lines/arrows are visibly aligned and connected as intended;
- nested containment is visually clear;
- the overlap example visibly places the covering box above the reconstructed lower box;
- no spurious hidden edges are visible through the covering box;
- text is legible and remains aligned plausibly to the authored grid;
- connectors crossing containing boxes are not accidentally clipped;
- explicit and auto-detected diagrams appear at the expected natural/direct locations;
- no raw source diagram is emitted in place of the visual;
- no repair/conversion warning requiring intervention occurs;
- no obvious diagram corruption attributable to Word versus LibreOffice conversion is present.

Reviewer role: human project owner/maintainer.

Waiver policy: no waiver within M0014.

Milestone completion command:

```powershell
./eng/review-check.ps1 --milestone M0014
```

Implementation must not fabricate approval. If all automated work is complete but approval is absent, terminate `AWAITING HUMAN REVIEW`.

## Documentation Policy

Implementation updates direct feature documentation required by DOC-01.

Broad documentation normalization remains a separate documentation-sync pass.

No guide documents are copied into the repository.

## Constrained Runtime

Tier 0-2 must not require external diagram tools or network access merely to generate text diagrams.

Tier 3 uses already-authorized installed Word/LibreOffice runtimes.

No new external content producer, package manager, browser, or network service is required by the product feature.

## Completion Expectations

Implementation owns milestone closure:

```text
read milestone + planning-seeded ledger
-> verify exact obligation/evidence-case set equality
-> decompose into bounded work packages
-> implement/validate/update ledger
-> run Tier 0-2
-> run real Word/LibreOffice Tier 3
-> produce human-review artifacts
-> obtain/record HR-M0014-01
-> freshly reread milestone
-> reconcile obligations/evidence cases <-> ledger <-> repository/evidence
-> completion audit
-> write compact durable completion evidence below
```

Passing tests alone does not complete M0014.

## Completion Evidence

Implementation populates this section only after final reconciliation/completion audit.

| Obligation / evidence case | Evidence | Validation gate / target | Result |
|---|---|---|---|
| pending | pending | pending | pending |

## Baseline-Executability Audit

Planning confirms that architecture, semantics, compatibility, scope, auto-detection policy, anonymous identity behavior, SVG boundary, layering/occlusion semantics, validation topology, and human-review policy are settled.

The expected implementation volume is decomposable into bounded work packages around parsing/geometry, auto-detection/config, SVG/semantic integration, DOCX authoring, and validation.

Remaining decisions are local implementation mechanics and do not require project-level invention.

M0014 is ready for GPT-5.6 Luna.

## Escalation Boundary

Return to planning for a material change to:

- source-as-layout semantics;
- explicit `diagram` syntax or stable-ID requirement;
- auto-detection default/candidate scope;
- anonymous auto-detected identity semantics;
- graph-layout or routing responsibility;
- supported geometry beyond the bounded orthogonal model;
- overlap/occlusion reconstruction policy;
- SVG product boundary or general source SVG support;
- existing Mermaid producer semantics;
- figure placement/reference compatibility;
- validation locus/equivalence;
- blocking human-review policy.

Implementation owns concrete code/test organization, thresholds, parser data structures, SVG mechanics, OpenXML SVG embedding details, diagnostics, and bounded execution sequencing within this contract.
