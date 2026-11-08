# M0018 — Template-based editable PowerPoint presentations

## Execution Profile

| Field | Value |
|---|---|
| Lifecycle state | ready |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna (project default) |
| Baseline executor readiness | confirmed for settled contract |
| Decision preservation | confirmed |
| Execution tractability | bounded by stable obligations and seeded ledger |
| Execution ledger | `.execution/M0018.md` |
| Scope size | large coherent feature |
| Implementation autonomy | high within fixed semantic contract |
| Documentation sync | separate pass, hint supplied |
| Focused validation | implementation-owned focused checks; documented assertions below |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real LibreOffice Impress opens output PPTX on authorized Windows interactive host; script/command provided as part of implementation |
| Validation locus/platform | Windows local, .NET 10 SDK, LibreOffice Impress installed for Tier 3; Office-independent authoring for Tier 1/2 |
| Environment authority | existing `docs/ENGINEERING.md`; install/host provisioning not authorized by this milestone |
| Consumer validation | build/package current dotnet tool, install to isolated local tool path and invoke PPTX scenario |
| Human review | blocking artifact-quality review HR-M0018-01 |

## Goal

Add template-first authoring of editable PowerPoint slides from readable Markdown using reusable prototype slides, fixed logical layouts, and typed text/image/table regions, without changing the established Word authoring contract.

## Target State

A user can configure a prepared PPTX template, author slides with level-1 headings and `{#layout}` selectors, route level-2 regions through `{#region}` selectors, and run YADG to generate ordered, editable PPTX slides. Ordinary prototype content, styling and geometry are preserved. Text overflow post-flight evaluation is explicitly absent.

## Scope and constraints

`docs/specs/PRESENTATION-AUTHORING.md` is the complete normative presentation contract: grammar, sixteen standard layouts, typed regions, plain visible prototype markers, rich-text and prepared-table prototype semantics, error behavior, CLI/workspace integration and strict exclusions. Preserve `docs/ARCHITECTURE.md` core/Office boundary, DOCX compatibility, existing path/asset security, style principles and explicit diagnostics. Implementation may choose internal types, libraries and editing algorithms without changing this contract. No Mandated PowerPoint installation or unmanaged SDK dependency for authoring.

## M0018 public configuration decision

The optional schema-v1 `presentations` root is a non-empty list of named decks with required `id`, `template`, ordered `sources`, and unique `output` basename, with no one-deck cap. Sources/templates may be shared across decks; the union of presentation sources is excluded from DOCX discovery. All configured PPTX outputs are committed as one independent, complete PPTX result set only after full PPTX success. The previous singular `presentation` key is rejected. In addition, `yadg presentation check` and `yadg presentation build` are supported direct/configuration-free commands with explicit file arguments and independent output behavior. Normative syntax, file/asset paths, validation and transaction semantics are in `docs/specs/PRESENTATION-AUTHORING.md` and `docs/specs/WORKSPACE-VALUES.md`. This decision is owned by planning, not delegated to implementation.

## Non-goals

No renderer-driven post-flight or overflow validation, UNO measurement, font shrinking, SmartArt/charts/animations/transitions, arbitrary layout DSL, HTML slides, custom layout catalog, speaker notes, diagram generation, reference/numbering semantics, table cell merge support, or broad DOCX refactoring. No guide-system copy or implementation code in planning overlay.

## Required authority

- `docs/specs/PRESENTATION-AUTHORING.md` (new; primary product contract)
- `docs/specs/WORKSPACE-VALUES.md` (schema-v1 workspace configuration authority)
- `docs/SPECS.md`, `docs/ARCHITECTURE.md`, `docs/ENGINEERING.md`, `docs/TERMINOLOGY.md`
- `docs/specs/AUTHORING-EXPERIENCE.md`, `docs/specs/STRUCTURED-CONTENT.md`, `docs/specs/TEMPLATE-PRESENTATION.md`
- `AGENTS.md`

## Acceptance criteria

- **AC-01** — A presentation source is divided into ordered slides by level-1 headings, defaulting to `content`; repeated `{#layout}` selectors correctly reuse a single prototype without unique-ID restrictions.
- **AC-02** — Level-2 `{#region}` headings route content without rendering their label, while level-3+ headings render as styled content; direct body content works for a single-region layout.
- **AC-03** — All sixteen built-in layout contracts and their region IDs/types are recognized and validate consistently; templates need only provide referenced layouts.
- **AC-04** — Plain visible slide/title/region tags discover unique editable prototypes, clone selected slides in Markdown order, remove control tags, preserve unrelated static elements and avoid emitting unselected prototypes.
- **AC-05** — Text regions render supported paragraphs, formatting, headings and nested lists as native editable PowerPoint text using prototype-owned presentation examples.
- **AC-06** — Image regions accept one local PNG/JPEG Markdown image, generate a native editable picture and preserve aspect ratio within prototype bounds.
- **AC-07** — Table regions accept one Markdown pipe table and populate a native editable table from prepared header/body row examples, preserving widths and formatting.
- **AC-08** — Pre-flight rejects malformed/unknown/missing/duplicate selectors and tags, incompatible region payloads, invalid assets, and invalid/unrepresentable table structures with precise source/slide/region diagnostics and no partial output mutation.
- **AC-09** — Optional schema-v1 `YADG.md` `presentations` supports one or multiple named declarations with required `id`, `template`, ordered `sources`, unique `output`, allowable reused source/template paths across decks, union-based DOCX discovery isolation, all-decks `check`/`build` behavior and independent atomic whole-set `YadgPresentations/` replacement. The singular mapping fails explicitly; existing DOCX/render/publish semantics remain intact.
- **AC-10** — Resulting slides are fully editable native PPTX objects, have retained relevant master/theme/media links, and open without repair in real LibreOffice Impress.
- **AC-11** — Focused positive/negative fixtures and ordinary repository validation exercise the supported presentation contracts, and realistic template fixture provenance is recorded with application/version/hash.
- **AC-12** — Current packaged `dotnet tool` is installed and invoked in isolation for an end-to-end presentation sample, with no reliance on a stale globally installed tool.
- **AC-13** — Standalone `yadg presentation check` and `yadg presentation build` accept explicit template and repeatable ordered source arguments without opening or requiring `YADG.md`; build writes only its explicit output transactionally; direct commands do not invoke DOCX, producers, render or publish and reject malformed CLI options or paths.
- **DOC-01** — Direct user-facing Markdown/template authoring documentation and a minimal working sample explain the full standard layout set, tags, region types, image/table examples, multiple-deck configuration, direct configuration-free commands, diagnostics and known overflow limitation.
- **REV-01** — HR-M0018-01 is approved by a human based on real Impress rendering and editability examination of representative generated deck(s).

## Selective evidence cases

| ID | Parent | Distinct proof |
|---|---|---|
| EC-04a | AC-04 | Repeated prototype cloning and slide order |
| EC-04b | AC-04 | Static shapes, master/theme, no leaked unselected prototypes or tags |
| EC-05a | AC-05 | Plain paragraphs and content headings |
| EC-05b | AC-05 | Nested bullets and mixed inline formatting |
| EC-08a | AC-08 | Invalid source selectors/region payloads |
| EC-08b | AC-08 | Invalid template marker/region/asset/table structures |
| EC-09a | AC-09 | Multiple named configured decks, shared source/template, successful whole-set replacement |
| EC-09b | AC-09 | Existing DOCX workspace regression, union discovery exclusion and output isolation |
| EC-09c | AC-09 | One invalid deck prevents any PPTX result-set replacement; invalid duplicate declaration/output and singular root diagnosed |
| EC-13a | AC-13 | Direct check with no YADG.md and malformed YADG.md present, without output mutation |
| EC-13b | AC-13 | Direct build with explicit output; output preservation on failure and isolation from configured decks |
| EC-10a | AC-10 | OOXML/native shape structural inspection |
| EC-10b | AC-10 | Real Impress open and human-observable editability |

## Validation gates

| Gate | Tier / target | Locus and prerequisites | Command / check | Proves | Evidence |
|---|---|---|---|---|---|
| VAL-01 | Tier 1 focused, presentation parser / Open XML package | Windows local, .NET 10 | Focused presentation tests selected from repository test project; executor records exact command | AC-01, AC-02, AC-03, EC-04a, EC-04b, EC-05a, EC-05b, AC-06, AC-07, EC-08a, EC-08b, EC-09a, EC-09c, EC-13a, EC-13b | pass/fail log, fixture hash and assertions |
| VAL-02 | Tier 2 standard repository | Windows local with existing SDK | `./eng/validate.ps1` | AC-11, EC-09b, DOC-01 | complete output/log and repository revision |
| VAL-03 | Tier 3 real Impress open | Windows interactive user with installed LibreOffice Impress; isolated profile, temporary copy; no new host installation | implementation adds documented `eng/test-m0018-tier3.ps1` and runs it | EC-10b, AC-10, AC-11 | application/version, input/output SHA-256, open result, screenshots or saved review artifacts, cleanup status |
| VAL-04 | Tier 4 packaged consumer | Windows local, .NET 10 SDK | `dotnet pack` current tool; `dotnet tool install --tool-path <temp-path> --add-source <current-pack-dir> Yadg` then invoke isolated shim for sample; version/package options adapted to repository | AC-12, AC-09, EC-09a, AC-13, EC-13b | isolated installed command and inspected resulting deck |
| VAL-05 | Tier 5 human | Windows human reviewer with Impress (PowerPoint optional) | `./eng/review-check.ps1 --milestone M0018` | REV-01, EC-10b | signed/recorded review outcome |

VAL-03 is acceptance only after a complete candidate exists and portable focused checks establish its expected semantics. VAL-04 runs on the current candidate package, not a previous release. Focused tests provide inexpensive feedback during implementation. If the real Impress runtime is missing, mark VAL-03 blocked (not passed), preserve available Tier 1/2/4 evidence and request an authorized host; do not install host-wide tools or create CI workflows without permission. A failed Impress open is a product failure rather than a missing-capability waiver. Do not claim automated layout/overflow fit.

## Human review

- Applicability: **blocking**, class: artifact-quality, canonical ID: **HR-M0018-01**, owner: M0018.
- Subject: at least one representative generated PPTX from an actual editable prototype made in PowerPoint or Impress, covering ordinary text, a multi-region slide, one image and one native table. Review static template preservation, slide order, formatting plausibility, absence of placeholders, editability of picture/table/text, no repair on open, and visible overflows as ordinary quality observations.
- Record real reviewer decision and evidence under existing `.review/` milestone conventions. Only `approved` satisfies REV-01; `rejected` returns to implementation. No automatic or fabricated approval, and no waiver in this milestone.
- Gate: `./eng/review-check.ps1 --milestone M0018` (extend the existing mechanism if required).

## Documentation policy and sync

Implementation owns direct authoritative changes to `docs/SPECS.md`, `docs/ARCHITECTURE.md`, `docs/TERMINOLOGY.md`, `docs/ENGINEERING.md` where needed for integration and explicit direct authoring documentation (`DOC-01`). Broader content cleanup is deferred through `.guide-sync/pending/M0018-presentations.md`. The new spec is authoritative immediately and is not research.

## Completion expectation and evidence

Implementation owns bounded work packages, evidence-case/obligation status and durable closure. Re-read this milestone and ledger at completion, reconcile exact obligation/case sets, and append compact evidence to this section before cleaning mutable ledger. Success requires current valid evidence for all required gates and human approval, not merely passing tests.

### Completion Evidence

*To be populated only by implementation after successful completion audit.*

## Escalation boundary

Return to planning for any needed change to layout catalog/IDs, Markdown grammar, PPTX template marker/typed-region semantics, external runtime validation or its accepted locus, output/CLI compatibility, or human review policy. Choose concrete implementation mechanics within this settled contract.
