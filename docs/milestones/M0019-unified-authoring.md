# M0019 — Unified authoring pipeline and cross-format consistency

## Execution Profile

| Field | Value |
|---|---|
| Lifecycle state | ready (after planning-ledger materialization at handoff; see ledger note) |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna (project default) |
| Baseline executor readiness | confirmed for settled semantic/validation contract |
| Decision preservation | confirmed |
| Execution tractability | large but one coherent cross-format contract; implementation groups into bounded resumable work packages |
| Execution ledger | `.execution/M0019.md` (planning seed transported with this branch under `docs/milestones/M0019-ledger-seed.md`; materialize locally before implementation; see note below) |
| Scope size | large coherent consolidation |
| Implementation autonomy | high within fixed public/runtime boundaries |
| Documentation sync | separate pass; only directly contradicted user docs updated in implementation |
| Focused validation | format-specific policy, values, producer, selected/declarative/direct CLI, strict/YOLO and transactions |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real pinned Bun/Mermaid CLI; real Word/Writer and PowerPoint/Impress application boundaries |
| Validation locus/platform | Windows local/interactive host as specified below; .NET SDK 10; installed desktop Office apps for Tier 3 |
| Environment authority | `docs/ENGINEERING.md` and `docs/specs/CONTENT-PRODUCERS.md`; existing authorized tools only; see validation constraints |
| Consumer validation | package and install current Yadg dotnet tool into isolated path; invoke DOCX and PPTX examples |
| Human review | blocking `HR-M0019-01`, exactly one finalized DOCX plus one PPTX |

## Goal

Integrate existing YADG features across DOCX and PPTX instead of letting presentation support become an isolated parallel pipeline: consistent Markdown policies, source values, Mermaid-derived images, strict/YOLO recovery, direct authoring, selected presentation workflows, observability, and validation. Preserve template-first formatting and format-specific structure.

## Target State

Both backends consume consistent common authoring semantics and the same configured external Mermaid producer boundary. Authors can build one DOCX or PPTX explicitly without a workspace; an optional explicit config exposes shared authoring features. Configured presentation-only/mixed workspaces and selected-deck commands behave predictably under strict and YOLO operation. Word/Writer and PowerPoint/Impress integration, pinned Bun invocation and installed-tool consumption prove the result. A human approves exactly one representative finalized DOCX and one representative PPTX.

## Scope and workstream boundaries

Planning identifies nine coherent concerns, **not an implementation-owned task breakdown**: (1) format-neutral policies/diagnostics/asset boundaries, (2) inline Markdown consistency, (3) external producer reuse, (4) strict/YOLO behavior, (5) direct DOCX and configuration-enabled direct commands, (6) values, (7) selected-deck CLI and inventory, (8) real and packed-tool verification, (9) direct user documentation and human review. The implementation agent owns its own bounded work packages, sequencing, refactors and test organization.

## Decisions and Constraints

`docs/specs/CROSS-FORMAT-AUTHORING.md` is normative for all new contracts. The following material decisions are settled:

- Existing `YADG.md` schema remains version 1, and existing workspaces must not require a migration. Optional direct `--config` is a separate YAML document, never an implicit workspace. Allowed keys and root version, parser behavior, path base and trust boundary are fixed in the spec.
- Retain separate Word and presentation Markdown structural models. Shared policy/asset/producer/diagnostic services do not impose common section/slide semantics, cross-format reference IDs or generic Office ASTs.
- Mermaid retains existing fences, mandatory stable ID, optional caption, configured external process contract and PNG product. PPTX `image` accepts one image or one Mermaid block; captions are metadata, not new visible numbered slide figures. The PPTX stable ID is unique within a deck.
- Template-owned roles govern PPTX inline code. The new distinct `{{codeInline}}` run exemplar is required in the relevant prototype scope only when styled code is actually rendered; old plain-code behavior is not silently claimed to be styled output. No hardcoded font fallback in strict mode.
- Values substitute literally in authored Markdown in both backends; PPTX additionally handles selected-prototype visible ordinary text, never template control markers or metadata; existing Word-template story behavior remains intact.
- YOLO recovery is explicit, bounded, visible and logged, not an indiscriminate catch/ignore. PPTX missing/failed image production yields a native visible region placeholder; unrepresentable slide/table structure and malformed configurations remain fatal.
- Direct document build creates a single **unfinalized authored DOCX**; it never invokes render/publish or touches workspace result directories. Direct PPTX retains existing one-file behavior, with explicit optional config and YOLO.
- `presentation check/build --id` is separate configured selection mode. It validates the global declaration table but does not validate or execute unrelated source/deck/producers. Selected build replaces one named output file; ordinary workspace build still replaces complete configured PPTX set. DOCX/PPTX result sets remain independently committed, not globally atomic.
- Authoring stays Office-independent. No direct use of COM/UNO in authoring services or external producer plugins. Preserve immutable real Office-origin committed compatibility fixtures; regenerate through the originating application if a fixture needs new prototype examples.
- No additional diagram producer kind, OpenXML generic renderer, custom PPTX layouts, post-flight overflow measurement, new semantic slide references, PDF producer output, container sandbox or arbitrary plugin API.

## Required authority

- `docs/specs/CROSS-FORMAT-AUTHORING.md` — M0019 primary public behavior and failure semantics.
- `docs/specs/PRESENTATION-AUTHORING.md` — PPTX prototype/layout/typed-region and existing CLI rules, except explicitly specialized.
- `docs/specs/CONTENT-PRODUCERS.md` — configured external process and Mermaid command/PNG contract.
- `docs/specs/AUTHORING-RESILIENCE.md` — strict/YOLO degradation and renderer fatal boundaries.
- `docs/specs/AUTHORING-EXPERIENCE.md` — Markdown policies, diagnostics and `check --list`.
- `docs/specs/WORKSPACE-VALUES.md` — schema-v1 values, reference and root configuration.
- `docs/specs/DOCUMENT-COMPOSITION.md`, `docs/specs/TEMPLATE-PRESENTATION.md`, `docs/specs/WORD-REFERENCES.md` — existing Word semantics and style ownership.
- `docs/ENGINEERING.md`, `docs/ARCHITECTURE.md`, `docs/SPECS.md`, `docs/TERMINOLOGY.md`, `AGENTS.md` — project architecture, environment and validation authority.
- Applicable guide scopes resolved during planning: repository-wide `base` and `artifact-first-runtime`, with concrete runtime/consumer specializations preserved here and in `docs/ENGINEERING.md`. Implementation does **not** need to consult the external guide.

## Acceptance Criteria

- **AC-01** — `markdown.codeInline` obeys `error`, `ignore` and `style` for source content in both DOCX and PPTX; PPTX styled code uses the required uniquely marked native exemplar and fails strictly when absent, preserving existing no-code templates.
- **AC-02** — `markdown.thematicBreak` obeys strict `error` and explicit `ignore` in both backends, including correct contextual diagnostics and omission without adjacent content loss.
- **AC-03** — PPTX supports nested/composed bold, italic and strikethrough in supported text/title/cell scopes with native editable formatting; unsupported container structures remain explicit errors.
- **AC-04** — PPTX supported external hyperlinks become native clickable links with unchanged visible words; unsupported link destinations preserve visible text without fabricated targets and report the specified warning.
- **AC-05** — Literal nonrecursive `{{value:id}}` substitutions operate on authored DOCX/PPTX text and selected PPTX prototype visible text across the named scopes, preserving surrounding/template formatting and control syntax.
- **AC-06** — Strict missing values fail with precise diagnostics and no output mutation; YOLO retains unresolved valid tokens visibly with degradation, while malformed tokens stay fatal.
- **AC-07** — Both output formats use the existing Mermaid external-process command/PNG product validation and isolation/cleanup semantics; checks do not create normal output, and builds consume current rather than stale generated data.
- **AC-08** — A PPTX image region accepts one properly identified Mermaid block, emits one native picture from its PNG into prototype bounds, and preserves deck-scoped ID uniqueness and caption-as-metadata semantics.
- **AC-09** — Producer startup/nonzero/timeout/invalid-image/missing-config failures yield strict actionable diagnostics and preserve existing DOCX/PPTX output sets with cleaned temporary assets.
- **AC-10** — PPTX YOLO missing in-bound images and failed/unavailable Mermaid producer recover through visibly identifiable native region placeholders with accurate single-event degradation reporting.
- **AC-11** — Strict/YOLO policy is consistent across configured and direct authoring; the enumerated recoveries succeed, and malformed config, ambiguous prototypes/IDs, structurally incompatible typed regions, escaping paths and no useful output remain fatal.
- **AC-12** — `yadg document check` validates one explicit DOCX template and ordered explicit sources without workspace marker/discovery and reports same semantic/template diagnostics as workspace DOCX check.
- **AC-13** — `yadg document build` produces one authored, unfinalized DOCX at explicit output atomically, using existing Word authoring semantics; failure preserves prior output and never invokes renderer/publisher or mutates workspace result sets.
- **AC-14** — Direct `presentation check/build` retain M0018 source ordering, validation and one-file transaction semantics while accepting optional explicit config/YOLO; a malformed ambient `YADG.md` is irrelevant to both direct command families.
- **AC-15** — Explicit direct `--config` parses only the declared schema-v1 authoring keys and enforces value types, executable trust/path base, forbidden roots and malformed-YAML errors; absent `--config` uses defaults and executes no ambient configured producers.
- **AC-16** — Direct `document` and `presentation` commands support invocation-local `--yolo` with truthful degradation counts and the same enumerated format-specific recovery/fatal boundaries as configured authoring.
- **AC-17** — `presentation check --id` and corresponding selected mode locate one configured deck, reject missing ID/global invalid declarations, and avoid checking DOCX or unrelated deck inputs and producers.
- **AC-18** — `presentation build --id` commits only the selected named PPTX output on success and preserves all other configured output bytes on success or failure; no-ID workspace `build` retains full-set replacement semantics.
- **AC-19** — `yadg check --list` deterministically reports declared presentations with template, ordered source paths and output alongside the existing DOCX/source/reference inventory, including discoverable entries when invalid.
- **AC-20** — DOCX output retains native template/paragraph/list/field/reference semantics under the new shared authoring path through both actual Word and LibreOffice Writer integrations, including real Mermaid and values.
- **AC-21** — PPTX output retains editable native slides, master/theme/prototype formatting, static shapes, picture/text/table objects under the new content path through actual PowerPoint and Impress integrations.
- **AC-22** — The current packed Yadg dotnet tool installs and executes representative direct DOCX, direct PPTX, explicit config, YOLO, and configured selection invocations from an isolated consumer without relying on a global/stale tool.
- **AC-23** — Existing DOCX-only features (grid tables, Word semantic references, blockquotes, fenced code, text diagrams, generated captions and renderer/publish contracts) remain functional; no new PPTX structures are fabricated to imitate them.
- **DOC-01** — Current authoring/CLI/YOLO/producer/PPTX/DOCX user docs and authoritative indexes show working examples, configuration-free vs explicit config vs workspace distinction, new template role examples, producer execution trust and documented format limits.
- **REV-01** — Human artifact-quality review `HR-M0019-01` approves the designated finalized DOCX and PPTX tied to current Tier 3 evidence and hashes.

## Acceptance Evidence Topology

Use selective distinct evidence paths; none of these is a work-package prescription.

| ID | Parent | Distinct evidence case | Why independent |
|---|---|---|---|
| EC-01a | AC-01 | DOCX codeInline error/ignore/style and template role | Word parser and style implementation |
| EC-01b | AC-01 | PPTX codeInline policies, per-scope exemplar and compatibility failure | Separate DrawingML template handling |
| EC-05a | AC-05 | Source-authored value in DOCX plus retained Word story substitution | Word semantic and template-story surfaces |
| EC-05b | AC-05 | PPTX source title/text/table and selected-prototype static text | Different PPTX scopes and control-marker exclusion |
| EC-07a | AC-07 | Real configured producer/validated PNG consumed by both format adapters | External command plus two authoring boundaries |
| EC-07b | AC-07 | Check-only temporary products, failure cleanup and output freshness | Lifecycle different from normal build |
| EC-08a | AC-08 | Configured PPTX Mermaid region | Workspace declaration/config path |
| EC-08b | AC-08 | Direct PPTX Mermaid region using explicit config | Standalone invocation/config isolation |
| EC-10a | AC-10 | Missing in-root local image placeholder | Asset-file failure path |
| EC-10b | AC-10 | Missing/failing Mermaid producer placeholder | External process failure path |
| EC-11a | AC-11 | YOLO permitted recovery and correct counters | Recovered path |
| EC-11b | AC-11 | YOLO rejects truly fatal semantic, structural, security errors | Fatal boundary remains strict |
| EC-13a | AC-13 | Direct DOCX output replaces only named file on success | One-file commit path |
| EC-13b | AC-13 | Direct DOCX failed build preserves prior bytes/inputs | Rollback path |
| EC-15a | AC-15 | Valid direct config values/producer/references/Markdown settings | Separate explicit parser/source |
| EC-15b | AC-15 | Missing/malformed/forbidden config, ambient workspace ignored | Trust and error boundary |
| EC-17a | AC-17 | One selected deck checked without loading unrelated DOCX/decks | Targeted selection path |
| EC-17b | AC-17 | Missing ID, global duplicate/output collision and mode conflict | CLI/schema failure path |
| EC-18a | AC-18 | Successful selected build changes only target, full workspace build still replaces set | Different commit scopes |
| EC-18b | AC-18 | Selected failure leaves all existing files unchanged | Targeted rollback |
| EC-20a | AC-20 | Word-origin real Word finalization with new source values/Mermaid | Word native runtime |
| EC-20b | AC-20 | LibreOffice-origin real Writer finalization with new source values/Mermaid | Independent Writer native runtime |
| EC-21a | AC-21 | PowerPoint-origin real PowerPoint generated-slide reopen/editability | PowerPoint native runtime |
| EC-21b | AC-21 | Impress-origin real Impress generated-slide reopen/editability | Independent Impress native runtime |
| EC-22a | AC-22 | Installed tool direct DOCX/config/YOLO use | Package consumer Word path |
| EC-22b | AC-22 | Installed tool direct PPTX/selected deck/config use | Package consumer PowerPoint path |

## Validation Gates

| ID | Depth / target | Locus, capability, permissions | Command / execution trigger | Proves | Evidence |
|---|---|---|---|---|---|
| VAL-01 | Tier 1 focused authoring and CLI integration | Local Windows .NET SDK 10, source build; no Office required | Executor records exact focused `dotnet test` filter(s), run as relevant subsystems land | AC-02, AC-03, AC-04, AC-06, AC-09, AC-12, AC-14, AC-16, AC-19, AC-23, EC-01a, EC-01b, EC-05a, EC-05b, EC-07b, EC-08a, EC-08b, EC-10a, EC-10b, EC-11a, EC-11b, EC-13a, EC-13b, EC-15a, EC-15b, EC-17a, EC-17b, EC-18a, EC-18b | Precise test/scenario names, failures/diagnostics, revision and fixture hashes where used |
| VAL-02 | Tier 2 canonical repository validation | Windows local, existing .NET 10 and build prerequisites | `./eng/validate.ps1` once relevant source/test/doc surface stable; revisit freshness at closure | AC-01, AC-02, AC-03, AC-04, AC-05, AC-06, AC-07, AC-08, AC-09, AC-10, AC-11, AC-12, AC-13, AC-14, AC-15, AC-16, AC-17, AC-18, AC-19, AC-23, DOC-01 (through focused cases above where applicable) | Build/test output, repository revision, cross-format regression, fixture-provenance/hygiene results |
| VAL-03 | Tier 3 external producer | Local Windows with pinned Bun and Mermaid CLI via `eng/test-tools.json`, existing authorized user-scope installation/resolve; trusted executable config | Executor adds documented `./eng/test-m0019-mermaid-tier3.ps1` or equivalent; run when producer service and both consumers are integrated | EC-07a, EC-07b, EC-08a, EC-08b, EC-10b, AC-09 | Actual executable/package versions, command/exit, PNG validation/hashes, source/output identities and cleanup, strict/YOLO failure receipts |
| VAL-04 | Tier 3 real DOCX runtime | Interactive Windows with activated desktop Word COM and installed LibreOffice Writer, established isolation/cleanup policy; use existing native-origin fixtures | Executor adds `./eng/test-m0019-docx-tier3.ps1` or equivalent; run after DOCX source values/producer/YOLO integration stabilizes | EC-20a, EC-20b, AC-05, AC-06, AC-23 | Template provenance hashes, authored/finalized DOCX hashes, renderer/app versions, fields/styles/figures/values checks and cleanup receipts |
| VAL-05 | Tier 3 real PPTX runtime | Interactive Windows with activated PowerPoint and installed Impress/UNO, isolated profiles and owned-process cleanup; existing native-origin PPTX fixtures | Executor adds `./eng/test-m0019-pptx-tier3.ps1` or equivalent; run after PPTX authoring feature set stabilizes | EC-21a, EC-21b, EC-01b, EC-05b, EC-08a, AC-03, AC-04 | Both application versions, source/template hashes, output hashes, native object/style/master/format comparisons, no repair, cleanup receipts |
| VAL-06 | Tier 4 installed current Yadg tool consumer | Windows .NET 10, isolated `dotnet tool install --tool-path` and package feed; no global tool dependency | Executor adds documented `./eng/test-m0019-tier4.ps1` or equivalent; run when public CLI and packaging are stable | EC-22a, EC-22b, AC-12, AC-14, AC-16, AC-17, AC-18 | Current package SHA, tool install path, command exits, output hashes, malformed/YOLO case |
| VAL-07 | Tier 5 human artifact review | Human on Windows with Word and Impress; follows both real runtime receipts | `./eng/review-check.ps1 --milestone M0019` after two designated artifacts and manifest produced and current | REV-01, DOC-01 | Genuine signed human decision bound to exactly two candidate paths, SHA-256, provenance and Tier 3 receipt identities |

VAL-03, VAL-04 and VAL-05 are **independent** acceptance stages; a missing runtime blocks its gate without invalidating current evidence for another. During implementation prefer VAL-01 focused feedback over full Tier 3/4 reruns. VAL-02 is the stable integrated repository gate, VAL-06 is the stable packaged CLI gate. At completion verify all gates are current against relevant source/config/provenance; rerun only stale/absent evidence, with conservative invalidation when dependency changes are uncertain. Full regeneration of human review candidates invalidates pending human decision until reviewed; once M0019 is complete, review is historical and not perpetually repeated.

### Execution, provisioning and infrastructure permissions

- **Execution permitted:** Windows-local SDK restore/build/tests, use of installed and authorized Word/PowerPoint/LibreOffice, invoke configured trusted Bun/Mermaid CLI, create isolated temporary profiles/output copies and isolated user-scope `dotnet tool` installs; no automated Office launch on noninteractive/service sessions.
- **Provisioning:** restore existing declared .NET dependencies and isolated package consumers. Do **not** install new persistent/system-wide Word/PowerPoint/LibreOffice/Bun/toolchains, modify registered Office configuration or provision privileged environments. User-scoped pinned Bun/package download is permitted only if already established in existing project test tooling on the authorized host; otherwise report unavailable tool capability.
- **Infrastructure:** do not create or modify CI workflows/runners, global permissions/secrets or remote services for this milestone. Accepted Tier 3 locus is an authorized local interactive Windows user session; a different existing interactive Windows host with the same native runtime capability is equivalent if recorded. No simulated Office app or OpenXML-only substitute proves Tier 3. Missing capability is `BLOCKED`, never a waiver. Distinguish missing runtime from failing product. No persistent cache/network sandbox guarantee is added to Mermaid.
- Run real Word automation serially; isolate LibreOffice profile and cleanup owned processes and temporary authoring products. Producer execution may run via configured Bun and may access network according to its command; source-controlled config is trusted executable configuration.

## Human Review

- Applicability: **blocking**. Class: artifact-quality. Canonical ID: **HR-M0019-01**. Owner: M0019.
- Review subjects: **exactly two named artifacts** under `artifacts/review/evidence/M0019/`: `finalized-word.docx` (Word-origin template, finalized by real Word) and `deck-impress.pptx` (Impress-origin prototype, generated by YADG, reopened in real Impress). These are one DOCX and one PPTX; no additional alternative decks/docs presented as human review targets.
- Bind both files' SHA-256, source/content/config and template fixture hashes, tool/package identity, application versions, and current VAL-03/04/05 receipts in `artifacts/review/evidence/M0019/manifest.json`. The reviewer examines Markdown-produced source values, inline code/strikethrough/links as applicable, Mermaid image and explicit YOLO placeholder or a visibly distinguished separate test slide, static template-owned formatting, editable text/table/picture structures, no leaked control markers, and no repair prompts. YOLO degradations must be visually distinguishable and traceable to manifest/logs; no requirement for post-flight overflow scanner.
- Reviewer: actual human with Word and Impress access. Only genuine `approved` bound to exact artifact hashes counts; `rejected` returns to implementation. No waiver and no agent-generated approval.
- Extend the existing milestone review-check mechanism to verify current artifact hashes and evidence; gate remains pending until signed review. Do not pre-create fictitious approval content.

## Documentation policy and deferred sync

Update directly contradicted public guides and CLI examples for this milestone, including Markdown, templates, presentation authoring, direct Word authoring, Mermaid, YOLO and direct config. Keep this spec authoritative. Broad README/catalog normalizations and release notes remain a later documentation-sync pass; a focused hint may be added under `.guide-sync/pending/`. No release-version bump or package publication is part of M0019.

## Execution-ledger handoff for direct-branch planning

Repository hygiene intentionally excludes `.execution/` and rejects tracked operational files. To avoid committing prohibited operational state while still preserving a **planning-authored seed**, this branch carries the exact transport text at `docs/milestones/M0019-ledger-seed.md`. The implementer must **materialize that content verbatim** as local, ignored `.execution/M0019.md` before any implementation; then remove the transport seed from the implementation PR before merge. This is a transport step, not a new execution contract. Until the ignored ledger is materialized, the work is not implementation-ready locally; the agent must not invent rows, statuses or proof mappings. The seed has no tasks, evidence, executed gates or completion claims. At handoff, verify obligation and evidence-case sets match exactly.

## Completion Expectations and Evidence

Implementation owns bounded work-package mapping, persistent ledger status, criterion/evidence-case-specific proof, affected-gate freshness and final reconciliation. Completion is the **exact equality** of (a) milestone obligation IDs and ledger obligations; (b) selective `EC-*` IDs and ledger cases; (c) current validation receipts and required gates; and (d) real human approval, followed by the completion audit. Persist a compact durable result table here **after** closure; do not write fictitious completion evidence during planning. The ignored ledger may then be removed per repository policy.

## Escalation boundary

Implementation chooses classes, tests, algorithms, and staging mechanics consistent with this contract. Return to planning if a material alternative affects schema, semantic parity, YOLO recovery/fatal boundary, producer trust, CLI mode distinctions, transaction ownership, permitted runtime locus/provisioning, human review, or evidence requirements. Do not read external guide files to reconstruct project rules.
