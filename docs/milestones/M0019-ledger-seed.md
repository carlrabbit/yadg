# M0019 — Planning-seeded execution ledger

> TRANSPORT-ONLY PLANNING SEED: Copy the complete contents of this file verbatim to ignored `.execution/M0019.md` on the implementation checkout before making implementation changes. The repository rejects tracked operational state. Delete this transport file from the implementation PR after materialization. The authoritative contract is `docs/milestones/M0019-unified-authoring.md` and its referenced project specs; this seed has no implementation status or evidence.

## Milestone authority

- Primary milestone: `docs/milestones/M0019-unified-authoring.md`
- Execution profile: ai-executed-human-reviewed
- Initial state: planning-seeded; all obligations, cases and validation gates pending
- Implementation owns bounded work packages, implementation mapping/evidence, validation receipts, resumable status, and closure audit

## Obligation coverage — planning seed

| Obligation ID | Milestone obligation (verbatim) | Initial state | Implementation evidence |
|---|---|---|---|
| AC-01 | `markdown.codeInline` obeys `error`, `ignore` and `style` for source content in both DOCX and PPTX; PPTX styled code uses the required uniquely marked native exemplar and fails strictly when absent, preserving existing no-code templates. | pending | — |
| AC-02 | `markdown.thematicBreak` obeys strict `error` and explicit `ignore` in both backends, including correct contextual diagnostics and omission without adjacent content loss. | pending | — |
| AC-03 | PPTX supports nested/composed bold, italic and strikethrough in supported text/title/cell scopes with native editable formatting; unsupported container structures remain explicit errors. | pending | — |
| AC-04 | PPTX supported external hyperlinks become native clickable links with unchanged visible words; unsupported link destinations preserve visible text without fabricated targets and report the specified warning. | pending | — |
| AC-05 | Literal nonrecursive `{{value:id}}` substitutions operate on authored DOCX/PPTX text and selected PPTX prototype visible text across the named scopes, preserving surrounding/template formatting and control syntax. | pending | — |
| AC-06 | Strict missing values fail with precise diagnostics and no output mutation; YOLO retains unresolved valid tokens visibly with degradation, while malformed tokens stay fatal. | pending | — |
| AC-07 | Both output formats use the existing Mermaid external-process command/PNG product validation and isolation/cleanup semantics; checks do not create normal output, and builds consume current rather than stale generated data. | pending | — |
| AC-08 | A PPTX image region accepts one properly identified Mermaid block, emits one native picture from its PNG into prototype bounds, and preserves deck-scoped ID uniqueness and caption-as-metadata semantics. | pending | — |
| AC-09 | Producer startup/nonzero/timeout/invalid-image/missing-config failures yield strict actionable diagnostics and preserve existing DOCX/PPTX output sets with cleaned temporary assets. | pending | — |
| AC-10 | PPTX YOLO missing in-bound images and failed/unavailable Mermaid producer recover through visibly identifiable native region placeholders with accurate single-event degradation reporting. | pending | — |
| AC-11 | Strict/YOLO policy is consistent across configured and direct authoring; the enumerated recoveries succeed, and malformed config, ambiguous prototypes/IDs, structurally incompatible typed regions, escaping paths and no useful output remain fatal. | pending | — |
| AC-12 | `yadg document check` validates one explicit DOCX template and ordered explicit sources without workspace marker/discovery and reports same semantic/template diagnostics as workspace DOCX check. | pending | — |
| AC-13 | `yadg document build` produces one authored, unfinalized DOCX at explicit output atomically, using existing Word authoring semantics; failure preserves prior output and never invokes renderer/publisher or mutates workspace result sets. | pending | — |
| AC-14 | Direct `presentation check/build` retain M0018 source ordering, validation and one-file transaction semantics while accepting optional explicit config/YOLO; a malformed ambient `YADG.md` is irrelevant to both direct command families. | pending | — |
| AC-15 | Explicit direct `--config` parses only the declared schema-v1 authoring keys and enforces value types, executable trust/path base, forbidden roots and malformed-YAML errors; absent `--config` uses defaults and executes no ambient configured producers. | pending | — |
| AC-16 | Direct `document` and `presentation` commands support invocation-local `--yolo` with truthful degradation counts and the same enumerated format-specific recovery/fatal boundaries as configured authoring. | pending | — |
| AC-17 | `presentation check --id` and corresponding selected mode locate one configured deck, reject missing ID/global invalid declarations, and avoid checking DOCX or unrelated deck inputs and producers. | pending | — |
| AC-18 | `presentation build --id` commits only the selected named PPTX output on success and preserves all other configured output bytes on success or failure; no-ID workspace `build` retains full-set replacement semantics. | pending | — |
| AC-19 | `yadg check --list` deterministically reports declared presentations with template, ordered source paths and output alongside the existing DOCX/source/reference inventory, including discoverable entries when invalid. | pending | — |
| AC-20 | DOCX output retains native template/paragraph/list/field/reference semantics under the new shared authoring path through both actual Word and LibreOffice Writer integrations, including real Mermaid and values. | pending | — |
| AC-21 | PPTX output retains editable native slides, master/theme/prototype formatting, static shapes, picture/text/table objects under the new content path through actual PowerPoint and Impress integrations. | pending | — |
| AC-22 | The current packed Yadg dotnet tool installs and executes representative direct DOCX, direct PPTX, explicit config, YOLO, and configured selection invocations from an isolated consumer without relying on a global/stale tool. | pending | — |
| AC-23 | Existing DOCX-only features (grid tables, Word semantic references, blockquotes, fenced code, text diagrams, generated captions and renderer/publish contracts) remain functional; no new PPTX structures are fabricated to imitate them. | pending | — |
| DOC-01 | Current authoring/CLI/YOLO/producer/PPTX/DOCX user docs and authoritative indexes show working examples, configuration-free vs explicit config vs workspace distinction, new template role examples, producer execution trust and documented format limits. | pending | — |
| REV-01 | Human artifact-quality review `HR-M0019-01` approves the designated finalized DOCX and PPTX tied to current Tier 3 evidence and hashes. | pending | — |

## Evidence-case coverage — planning seed

| Evidence case ID | Parent obligation | Distinct required proof | Initial state | Implementation evidence |
|---|---|---|---|---|
| EC-01a | AC-01 | DOCX codeInline error/ignore/style and template role | pending | — |
| EC-01b | AC-01 | PPTX codeInline policies, per-scope exemplar and compatibility failure | pending | — |
| EC-05a | AC-05 | Source-authored value in DOCX plus retained Word story substitution | pending | — |
| EC-05b | AC-05 | PPTX source title/text/table and selected-prototype static text | pending | — |
| EC-07a | AC-07 | Real configured producer/validated PNG consumed by both format adapters | pending | — |
| EC-07b | AC-07 | Check-only temporary products, failure cleanup and output freshness | pending | — |
| EC-08a | AC-08 | Configured PPTX Mermaid region | pending | — |
| EC-08b | AC-08 | Direct PPTX Mermaid region using explicit config | pending | — |
| EC-10a | AC-10 | Missing in-root local image placeholder | pending | — |
| EC-10b | AC-10 | Missing/failing Mermaid producer placeholder | pending | — |
| EC-11a | AC-11 | YOLO permitted recovery and correct counters | pending | — |
| EC-11b | AC-11 | YOLO rejects truly fatal semantic, structural, security errors | pending | — |
| EC-13a | AC-13 | Direct DOCX output replaces only named file on success | pending | — |
| EC-13b | AC-13 | Direct DOCX failed build preserves prior bytes/inputs | pending | — |
| EC-15a | AC-15 | Valid direct config values/producer/references/Markdown settings | pending | — |
| EC-15b | AC-15 | Missing/malformed/forbidden config, ambient workspace ignored | pending | — |
| EC-17a | AC-17 | One selected deck checked without loading unrelated DOCX/decks | pending | — |
| EC-17b | AC-17 | Missing ID, global duplicate/output collision and mode conflict | pending | — |
| EC-18a | AC-18 | Successful selected build changes only target, full workspace build still replaces set | pending | — |
| EC-18b | AC-18 | Selected failure leaves all existing files unchanged | pending | — |
| EC-20a | AC-20 | Word-origin real Word finalization with new source values/Mermaid | pending | — |
| EC-20b | AC-20 | LibreOffice-origin real Writer finalization with new source values/Mermaid | pending | — |
| EC-21a | AC-21 | PowerPoint-origin real PowerPoint generated-slide reopen/editability | pending | — |
| EC-21b | AC-21 | Impress-origin real Impress generated-slide reopen/editability | pending | — |
| EC-22a | AC-22 | Installed tool direct DOCX/config/YOLO use | pending | — |
| EC-22b | AC-22 | Installed tool direct PPTX/selected deck/config use | pending | — |

## Validation gates — planning seed

| Gate ID | Depth / target | Required locus/capabilities | Target invocation / trigger | Proves (contract IDs) | Initial state |
|---|---|---|---|---|---|
| VAL-01 | Tier 1 focused authoring and CLI integration | Local Windows .NET SDK 10, source build; no Office required | Executor records exact focused `dotnet test` filter(s), run as relevant subsystems land | AC-02, AC-03, AC-04, AC-06, AC-09, AC-12, AC-14, AC-16, AC-19, AC-23, EC-01a, EC-01b, EC-05a, EC-05b, EC-07b, EC-08a, EC-08b, EC-10a, EC-10b, EC-11a, EC-11b, EC-13a, EC-13b, EC-15a, EC-15b, EC-17a, EC-17b, EC-18a, EC-18b | pending |
| VAL-02 | Tier 2 canonical repository validation | Windows local, existing .NET 10 and build prerequisites | `./eng/validate.ps1` once relevant source/test/doc surface stable; revisit freshness at closure | AC-01, AC-02, AC-03, AC-04, AC-05, AC-06, AC-07, AC-08, AC-09, AC-10, AC-11, AC-12, AC-13, AC-14, AC-15, AC-16, AC-17, AC-18, AC-19, AC-23, DOC-01 (through focused cases above where applicable) | pending |
| VAL-03 | Tier 3 external producer | Local Windows with pinned Bun and Mermaid CLI via `eng/test-tools.json`, existing authorized user-scope installation/resolve; trusted executable config | Executor adds documented `./eng/test-m0019-mermaid-tier3.ps1` or equivalent; run when producer service and both consumers are integrated | EC-07a, EC-07b, EC-08a, EC-08b, EC-10b, AC-09 | pending |
| VAL-04 | Tier 3 real DOCX runtime | Interactive Windows with activated desktop Word COM and installed LibreOffice Writer, established isolation/cleanup policy; use existing native-origin fixtures | Executor adds `./eng/test-m0019-docx-tier3.ps1` or equivalent; run after DOCX source values/producer/YOLO integration stabilizes | EC-20a, EC-20b, AC-05, AC-06, AC-23 | pending |
| VAL-05 | Tier 3 real PPTX runtime | Interactive Windows with activated PowerPoint and installed Impress/UNO, isolated profiles and owned-process cleanup; existing native-origin PPTX fixtures | Executor adds `./eng/test-m0019-pptx-tier3.ps1` or equivalent; run after PPTX authoring feature set stabilizes | EC-21a, EC-21b, EC-01b, EC-05b, EC-08a, AC-03, AC-04 | pending |
| VAL-06 | Tier 4 installed current Yadg tool consumer | Windows .NET 10, isolated `dotnet tool install --tool-path` and package feed; no global tool dependency | Executor adds documented `./eng/test-m0019-tier4.ps1` or equivalent; run when public CLI and packaging are stable | EC-22a, EC-22b, AC-12, AC-14, AC-16, AC-17, AC-18 | pending |
| VAL-07 | Tier 5 human artifact review | Human on Windows with Word and Impress; follows both real runtime receipts | `./eng/review-check.ps1 --milestone M0019` after two designated artifacts and manifest produced and current | REV-01, DOC-01 | pending |

## Implementation-owned state (empty at planning)

Implementation adds work packages, obligation-to-package mappings, receipts, durable resume state and completion audit here. Do not treat this seed as completed work. A real human review remains independently blocking.
