# M0008 execution ledger — external content producers and Mermaid

Milestone: `M0008-external-content-producers-mermaid`
Status: complete

## Work packages

1. Workspace producer configuration and process boundary
   - Extend schema-v1 front matter with validated `producers.mermaid`.
   - Add bounded direct-process execution, temporary input/output ownership, diagnostics, and PNG validation.
2. Mermaid semantic parsing and integration
   - Parse the constrained fenced syntax into generated semantic figures.
   - Preserve global IDs, natural/direct placement, caption, and reference semantics through existing figure code.
3. CLI lifecycle and documentation
   - Render producers during load/check/build before authored-output mutation.
   - Clean derived files and document package-manager-neutral configuration.
4. Validation and completion audit
   - Add focused integration coverage, run Tier 0–2 and the mandatory manifest-driven Tier 3 Bun/Mermaid/DOCX path.
   - Freshly reread authority and map every acceptance criterion to repository/evidence.

## Acceptance evidence map

Evidence is filled as each work package and validation tier completes.

| Acceptance area | Implementation evidence | Validation evidence |
|---|---|---|
| Workspace producer configuration | `WorkspaceValues`, producer parser | M0008 focused tests; Tier 0 |
| Markdown/semantic model | fenced Mermaid conversion and `YadgFigure` source | focused parser/semantic tests |
| Process runner | direct `ProcessStartInfo`, bounded wait, temp lifecycle, PNG checks | fake-producer tests; Tier 1 |
| Check/build | producer execution before Word authoring and cleanup | CLI integration; Tier 1/2/3 |
| Figure integration | existing `WordAuthoring` figure path | DOCX structural assertions; Tier 3 |
| Runtime neutrality/security | executable + argument prefix, no install code | focused config/process tests; review |
| Test-tool manifest | exact `eng/test-tools.json` consumption in Tier 3 harness | Tier 0 manifest validation; Tier 3 provenance |
| Tier-3 real target | pinned Bun download/version check/`bun x --bun --package` | mandatory Tier 3 runtime evidence |
| Compatibility/documentation | README/direct usage updates | `eng/validate.ps1`; completion audit |

## Validation log

- [x] Tier 0 — release build succeeded with zero errors; manifest was parsed by the Tier 3 test and checked for exact pins.
- [x] Tier 1 — focused M0008 tests passed: deterministic producer, missing configuration/output preservation, reserved flags, startup/nonzero/invalid-output diagnostics, PNG/DOCX figure path.
- [x] Tier 2 — `./eng/validate.ps1` passed: 36 passed, 0 failed, no external download.
- [x] Tier 3 — `eng/test-m0008-tier3.ps1` downloaded/verified Bun `1.4.2`, passed the exact Mermaid one-shot probe and real YADG check/build/DOCX test.
- [x] Fresh authority reread and completion audit completed after the final Tier 3 run.

## Criterion-level evidence

Each acceptance criterion is mapped below; shared evidence identifiers refer to the concrete files/tests and validation runs above.

- Workspace producer configuration: `WorkspaceValuesParser` accepts existing values-only workspaces and exact `producers.mermaid` executable/string arguments; it rejects unknown root/kind/key forms and reserved flags, requires an executable, permits unused configuration, and never requires Bun/npm. Evidence: `M0008IntegrationTests` focused run and Tier 0 build.
- Markdown/semantic model: `MarkdownDocumentParser` recognizes exact lowercase Mermaid fences, stable IDs, plain captions, nonempty source, rejects unknown/missing attributes and non-Mermaid fences, and creates `YadgFigure.GeneratedSource`; global merge keeps duplicate IDs invalid and generated source out of text. Evidence: parser path exercised by focused tests and real Tier 3.
- Process runner: `ExternalContentProducers.cs` resolves absolute/relative/PATH executables, uses `ProcessStartInfo.ArgumentList` without a shell, appends owned `-i`/`-o`, isolates per-invocation `.mmd`/`.png`, bounds waits, captures bounded stderr, diagnoses startup/timeout/exit/output failures, validates PNG signature/dimensions, and tracks best-effort cleanup. Evidence: deterministic producer test plus source inspection and Tier 3 command result.
- Check/build: `WorkspaceLoader` renders before template analysis/output creation; CLI cleanup runs in `finally`; generated output is never reused as stale source. Evidence: focused `check`/`build` test, preserved sentinel test, Tier 3 real check/build.
- Figure integration: generated figures reuse `YadgFigure` and existing `WordAuthoring` placement, intrinsic sizing, captions, bookmarks/REF validation, and direct/natural selection. Evidence: focused DOCX image/caption assertions and Tier 3 authored DOCX.
- Runtime neutrality/security: only trusted configured executable/arguments are run; no package installation or OOXML/plugin lifecycle API was added. Evidence: `ExternalContentProducers.cs`, README, and portable validation.
- Test-tool manifest: `eng/test-tools.json` is exact and is read by the Tier 3 test rather than duplicating package/version literals in test execution logic. Evidence: Tier 0/Tier 3.
- Tier-3 real target: `eng/test-m0008-tier3.ps1` downloads official Windows x64 Bun into test-owned temp state, verifies `1.4.2`, runs `bun x --bun --package @mermaid-js/mermaid-cli@11.17.0 mmdc`, and invokes real YADG. Evidence: final Tier 3 run passed; provenance printed platform, Bun, package/version, and probe path.
- Compatibility/documentation: README shows the fenced syntax and package-manager-neutral configuration; `eng/validate.ps1` remains download-free; all prior tests remain green; fixture content is synthetic; this ledger provides the mapping. Evidence: final Tier 2/Tier 3 runs and repository diff.

## Individual acceptance-criterion matrix

The following rows are intentionally one row per acceptance criterion, closing the earlier area-level-only mapping.

| Area | Criterion | Implementation | Evidence |
|---|---|---|---|
| Config | Existing values-only workspaces remain valid | `WorkspaceValuesParser` keeps optional `producers` | Tier 2, M0006 suite |
| Config | Nonempty executable and optional string arguments | producer schema parser | M0008 config test, Tier 0 |
| Config | Unknown root/kind keys fail | root/producer key validation | parser tests/source |
| Config | Missing Mermaid config fails when used | `YADG-PRODUCER-007` | missing-config test |
| Config | Unused configured producer is allowed | rendering only iterates generated figures | config path, Tier 2 |
| Config | Bun/npm are not product requirements | executable + argument model | README, Tier 3 configured Bun only |
| Markdown | Mermaid fence becomes figure at source anchor | fenced parser + `YadgFigure` | M0008 tests |
| Markdown | Stable ID mandatory/global | metadata regex + merge registry | parser/source, Tier 2 |
| Markdown | Duplicate IDs fail | existing `AddGlobal` registry | existing duplicate-ID tests |
| Markdown | Plain caption supported | caption capture into `AltText` | Tier 3 caption assertion |
| Markdown | Unknown attrs/empty source fail | info regex and source check | parser implementation, Tier 0 |
| Markdown | Other code fences remain unsupported | non-Mermaid fenced diagnostic | parser implementation |
| Markdown | Mermaid source is not document text | generated figure path only | focused/Tier 3 DOCX text assertions |
| Runner | Absolute/relative/PATH resolution | `ResolveExecutable` | runner implementation; Tier 1 fake command |
| Runner | No shell launch | `UseShellExecute=false`, `ArgumentList` | runner implementation |
| Runner | Arguments passed faithfully | `ArgumentList` loop | fake producer integration |
| Runner | YADG owns `-i`/`-o` | adapter appends both | reserved-flag test |
| Runner | Configured input/output flags rejected | parser reserved-flag diagnostic | reserved-flag test |
| Runner | Isolated temp files per invocation | GUID temp directory | focused build and cleanup |
| Runner | Timeout/start/nonzero failures diagnosed | bounded wait and diagnostics | startup/nonzero focused test; timeout bound in runner implementation |
| Runner | Useful stderr surfaced | bounded stderr capture | runner implementation |
| Runner | Missing/empty/non-PNG/undecodable output fails | structural PNG validation | producer implementation |
| Runner | Cleanup is best effort | CLI `finally` + tracked directories | focused temp hygiene |
| Runner | Stale output never accepted | fresh temp output per invocation | preserved-output test |
| Check/build | Check invokes producer without normal artifacts | load-time producer + no output creation | focused check |
| Check/build | Build validates before normal mutation | producer runs during load | preserved-output test, Tier 3 |
| Check/build | Invalid producer preserves prior output | failure before output authoring | sentinel test |
| Check/build | Successful build embeds PNG normally | existing `WordAuthoring` figure path | focused/Tier 3 DOCX |
| Check/build | No persistent image cache required | temp directory only | source inspection, Tier 2 |
| Figure | Natural placement works | existing section block traversal | focused/Tier 3 workspace |
| Figure | Direct placement suppresses natural emission | existing `direct` placement set | final Tier 3 direct tag |
| Figure | Intrinsic-size/fit behavior remains | existing `CreateFigureParagraph` | final Tier 3 DOCX |
| Figure | Caption uses existing behavior | `AltText` and caption prototype | final Tier 3 SEQ/caption |
| Figure | Captioned figure is numeric target | existing target/bookmark validation | final Tier 3 REF/bookmark |
| Figure | Uncaptioned figure is not numeric target | existing `BuildTargets` rules | existing M0004 tests |
| Runtime | Bun/npm/npx/mmdc/wrappers are valid config | no executable special-casing | schema/parser implementation |
| Runtime | Product never installs dependencies | no installer/cache code in product | source review, Tier 2 |
| Runtime | Producer config is trusted executable config | direct process boundary | README/spec alignment |
| Runtime | No OOXML/plugin lifecycle API exposed | producer lives in Core process adapter | source review |
| Manifest | Manifest parsed/validated | `JsonDocument` in Tier 3 test | Tier 3 run |
| Manifest | Versions come from manifest | test reads both pins | Tier 3 run |
| Manifest | Exact non-latest pins | `eng/test-tools.json` | manifest JSON validation |
| Tier 3 | Official Bun acquired locally | PowerShell test-owned cache | final Tier 3 provenance |
| Tier 3 | Bun version verified | `bun --version` equality check | final Tier 3 run |
| Tier 3 | Mermaid package/version from manifest | generated package argument | final Tier 3 run |
| Tier 3 | One-shot Bun execution with `--bun` | explicit `bun x --bun --package` | probe and YADG run |
| Tier 3 | No Node/npm/npx/mmdc dependency | only downloaded Bun invoked | script/source review |
| Tier 3 | Real Mermaid renders valid PNG | low-level probe + PNG checks | final Tier 3 run |
| Tier 3 | Real YADG check succeeds | Tier 3 test | final Tier 3 run |
| Tier 3 | Real YADG build produces image DOCX | Tier 3 test | image part assertion |
| Tier 3 | Direct/natural, caption/reference structures verified | direct tag, caption prototype, REF/SEQ/bookmarks | final Tier 3 assertions |
| Compatibility | `eng/validate.ps1` remains offline-capable | Tier 3 isolated from normal test | final Tier 2 run |
| Compatibility | M0002–M0007 remain green | unchanged suites | 36-test repository run |
| Compatibility | README/direct docs updated | Mermaid section in README | diff audit |
| Compatibility | Fixtures synthetic/non-confidential | generated test workspaces/content | test source audit |
| Compatibility | Ledger maps every criterion | this matrix | completion audit |

## Completion audit

- Milestone lifecycle is ready and all in-scope acceptance areas are implemented.
- M0007 prepared-table behavior remains intact; repository validation includes its existing tests.
- Overlay authority and required authority were freshly reread after implementation.
- No escalation-boundary decision was invented or required.
- Final status: complete.

## Decisions / escalations

No material product decision is currently required. Implementation choices remain within the milestone's stated boundary.
