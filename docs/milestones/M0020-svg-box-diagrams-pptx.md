# M0020 — Shared SVG box diagrams in PowerPoint and complete authoring guides

## Execution profile

| Field | Decision |
|---|---|
| Lifecycle | **ready** on application of this planning overlay and exact ledger seed |
| Execution | **ai-executed-human-reviewed** |
| Baseline model | GPT-5.6 Luna (project `.guide-profile.json` default) |
| Applied guide scopes | Repository-wide `base` + `artifact-first-runtime`, schema-v2 scope metadata; no additional local profile selected |
| Product role / maturity | `product-tool` / `initial-implementation` (project profile metadata; milestone adds no profile) |
| Planning/execution separation | Planning owns settled behavior, evidence and ledger coverage; implementation owns bounded work packages and mechanics |
| Planning-seeded ledger | `.execution/M0020.md` (included as transport in overlay; ignored by repository hygiene) |
| Canonical repository validation | `./eng/validate.ps1` on authorized Windows .NET 10 host |
| Real runtime gate | Native PowerPoint and LibreOffice Impress in authorized interactive Windows user session, origin-matched PPTX templates |
| Installed consumer gate | Current `.NET 10` Yadg package installed into isolated `dotnet tool --tool-path` |
| Human review | **blocking `HR-M0020-01`** for two SVG-containing decks: PowerPoint-origin in PowerPoint, Impress-origin in Impress |
| Documentation policy | In-milestone authored user documentation is a deliverable; unrelated documentation sync deferred |

## Goal and target state

Reuse YADG's existing M0014 character-grid parser, rectangle/occlusion model and deterministic **SVG** product across DOCX and PPTX. A `diagram` fence, or qualifying opt-in `text`/unlabeled fence, becomes an ordinary **vector SVG picture** inside an existing template-owned PPTX `image` region. YADG does **not** create native PowerPoint shape graphs. No new source language, template layout, schema key, global diagram producer or DOCX contract is introduced.

An author can follow the root README, bootstrap a workspace with `yadg init`, learn box diagrams from one short dedicated guide and configure the full supported workspace through one exhaustive user-facing `YADG.md` reference. Authoring and validation remain independent of Office automation; real PowerPoint and Impress acceptance evidence proves actual render compatibility.

## Settled decisions / non-goals

1. **Shared functionality:** Use the M0014 Office-independent parser/geometry/occlusion/SVG semantics without a second PPTX ASCII parser or format-specific layout engine. Separate DOCX section/figure and PPTX slide/region models remain intact. The existing M0019 shared policy/config/diagnostic and asset-lifecycle surfaces are reusable; no forced unified Office AST.
2. **SVG primary, no component shapes:** Embed real generated SVG as the PPTX vector picture product. A technically necessary companion raster *preview* is permitted only as an auxiliary interoperability fallback; a raster-only image fails. The picture itself is editable (move/resize/replace) but constituent diagram objects are not; no internal PowerPoint connector/box generation or bidirectional editing.
3. **Placement:** Extend only existing PPTX `image`-typed regions; exactly one image, identified Mermaid or text-grid diagram per region. Preserve contain-within-region aspect ratio, template layout and static artwork. No new layout families or mixed payload types.
4. **Source semantics:** Explicit `diagram {#stable-id caption="..."}` and optional `diagrams.autoDetect` use current grammar/heuristic. Explicit IDs deck-scoped; anonymous auto-detected image-region diagrams have no invented ID; captions are metadata, not slide-visible Word-style captions. No PPTX cross-slide references.
5. **Failure boundary:** Explicit malformed or invalid geometry/occlusion/ID, unsupported region usage, SVG packaging failure, and invalid template remain fatal even in YOLO, with atomic output preservation. Auto-detection keeps current conservative false-negative behavior. Diagram render does not invoke external Mermaid or silently rasterize. Preserve existing strict/YOLO and selected/direct/workspace authoring transactions.
6. **No new public config:** Existing schema-v1 `YADG.md` `diagrams.autoDetect` applies to presentation sources; direct `--config` uses its existing allowed schema, workspace paths and trust rules. `workspace.yaml` is **not** a YADG configuration filename. Preserve default false and all existing supported roots. No arbitrary imported `.svg` figures.
7. **Documentation is in scope:** New `docs/user/BOX-DIAGRAMS.md`, new consolidated `docs/user/CONFIGURATION.md`, updated root `README.md`, corrected public Markdown/presentation guide statements and navigational links. Examples must actually validate using current CLI and sample workspace. Do not create a third documentation authority that overrides specs; user guide explains them.
8. **Runtime boundary:** Keep authoring Office-independent. Accepted native validation loci are an existing authorized interactive Windows host with activated PowerPoint and installed LibreOffice Impress, or another existing equally capable authorized interactive Windows host recorded in evidence; no service automation, new CI runners/secrets, persistent machine-wide provisioning or external SVG renderer dependencies. A missing capability is BLOCKED, never silently replaced by OpenXML-only evidence.
9. **Excluded:** Native editable component shapes, per-shape styling, layout/reflow, new source graph semantics, connectors/routing, external `diagram` producer, SVG source file references, PDF/image exports as the product, template auto-generation, new runtime installers, broad documentation sync, release publication/version bump and generic post-flight overflow scanner.

## Required implementation-facing authority

- `docs/specs/PRESENTATION-TEXT-DIAGRAMS.md` — new normative SVG/PPTX placement and validation specialization **included in this overlay**.
- `docs/specs/TEXT-DIAGRAMS.md` — existing parser/geometry, occlusion, explicit/auto-detection and SVG contract.
- `docs/specs/PRESENTATION-AUTHORING.md` — existing prototype/image typed region, geometry and CLI.
- `docs/specs/CROSS-FORMAT-AUTHORING.md` — shared values/diagnostics/YOLO/producer/config/direct/selected modes.
- `docs/specs/WORKSPACE-VALUES.md` and `docs/specs/AUTHORING-EXPERIENCE.md` — YAML schema/bootstrap/Markdown policy.
- `docs/specs/CONTENT-PRODUCERS.md`, `docs/specs/AUTHORING-RESILIENCE.md` — preserve Mermaid external producer and fatal/recovery distinction.
- `docs/ARCHITECTURE.md`, `docs/ENGINEERING.md`, `docs/SPECS.md`, `AGENTS.md` — project boundaries and validation engineering.

The implementation agent does **not** need the external guide repository. Its ordinary authority is this ready milestone, listed project documents and current code/tests.

## Acceptance criteria and other obligations

- **AC-01** — A valid explicit identified box diagram in a PPTX image region is parsed by the existing shared geometry/occlusion implementation and yields the same deterministic SVG geometry/text as its DOCX counterpart, with no second parser or Office-specific layout logic.
- **AC-02** — PPTX embeds a genuine self-contained **SVG vector picture** with correct package relationships, image metadata and contain-bounded placement, never a raster-only replacement or individual PowerPoint shape graph; diagrams remain visually sharp at normal slide scaling.
- **AC-03** — Source grammar, mandatory explicit diagram ID, optional caption-as-metadata, deck-scoped uniqueness and region exclusivity are enforced; DOCX global-ID/reference behavior remains unchanged.
- **AC-04** — With `diagrams.autoDetect: true`, PPTX image regions accept qualifying unlabeled or `text` fences (identified or anonymous) through existing conservative heuristics; default false, weak candidates and unsupported contexts do not silently become diagrams.
- **AC-05** — The same diagram behavior works for configured whole-workspace, selected-deck and direct presentation modes, with explicit `--config` isolation, and no schema changes or Mermaid producer invocation.
- **AC-06** — Invalid explicit geometry/occlusion/attributes/IDs, duplicate deck IDs, wrong region type, and SVG/PPTX package failures give precise fatal diagnostics in strict and YOLO and preserve existing output bytes/other decks, without hidden substitutions.
- **AC-07** — Existing DOCX diagrams (identified and anonymous), PNG/JPEG, Mermaid and PPTX editable text/table/image behavior remain compatible; no changes to renderer/publisher or template layout catalog.
- **AC-08** — Native PowerPoint-origin→YADG→PowerPoint and Impress-origin→YADG→Impress outputs display identical authored vector diagram intent with correct labels, lines, nested boxes, arrows, occlusion and no clipping/overflow/repair; editable SVG pictures retain template-owned presentation structure.
- **AC-09** — One current packaged `.NET` tool installed in isolated consumer path successfully executes documented bootstrap/config and diagram PPTX authoring, not a stale global tool; installed output passes native runtime acceptance.
- **DOC-01** — `docs/user/BOX-DIAGRAMS.md` offers compact copy-ready examples for Unicode/ASCII, whitespace, nested/overlapping boxes, arrows, invalid cases, explicit/auto-detected fences and DOCX/PPTX semantics. Examples are executable/checked; no undocumented novel grammar.
- **DOC-02** — `docs/user/CONFIGURATION.md` is a **complete** annotated `YADG.md` schema-v1 user reference (all roots, child keys/types/defaults, restrictions, path/producer trust rules, direct `--config` limitations and examples); preserves canonical `YADG.md` filename and does not invent `workspace.yaml`.
- **DOC-03** — Existing `README.md` remains the landing page, links the two guides, gives a correct SVG diagram PPTX example and `yadg init` workflow; relevant Markdown/PowerPoint guides and indexes have no contradictory DOCX-only or obsolete configuration claims.
- **REV-01** — A genuine human approves exactly two designated current vector-diagram PPTX candidates after inspecting native-application renders and actual documents, bound to hashes/VAL-03/VAL-04 evidence through `HR-M0020-01`; no automatic or fabricated approval.

## Selective independent acceptance evidence

| Case | Parent | Independent proof |
|---|---|---|
| EC-01a | AC-01 | DOCX and PPTX share geometry/SVG for one explicit diagram, including nested/occluded shapes |
| EC-01b | AC-01 | Same SVG through multiple presentation slides/image regions with separate correct package parts |
| EC-02a | AC-02 | OOXML carries SVG vector payload, alt metadata and contain geometry (auxiliary bitmap, if any, cannot replace it) |
| EC-03a | AC-03 | Identified diagram metadata/ID/duplicate validation and typed region exclusivity |
| EC-04a | AC-04 | Explicit `diagrams.autoDetect: true`, anonymous and identified positive candidates |
| EC-04b | AC-04 | Default false, weak/ambiguous candidates, unsupported block/context stays invalid |
| EC-05a | AC-05 | Workspace and selected deck obey existing config and output ownership |
| EC-05b | AC-05 | Direct PPTX with vs without explicit config proves isolation |
| EC-06a | AC-06 | Strict/YOLO invalid explicit diagram preserves output |
| EC-06b | AC-06 | Packaging/region/duplicate failure preserves output and unrelated decks |
| EC-08a | AC-08 | Real PowerPoint-origin template→YADG→PowerPoint vector visual verification |
| EC-08b | AC-08 | Real Impress-origin template→YADG→Impress vector visual verification |
| EC-09a | AC-09 | Fresh isolated packed tool bootstrap and direct/config authoring |
| EC-09b | AC-09 | Installed tool selected/workspace authoring and real output acceptance |
| EC-10a | DOC-01 | Guide examples exercised against parser/CLI |
| EC-11a | DOC-02 | Complete schema-v1 coverage reconciled against current parser/specs, including disallowed direct roots |

## Validation gates

| Gate | Tier / target | Locus / capability and trigger | Proves | Required evidence |
|---|---|---|---|---|
| **VAL-01** | Tier 1 focused parser/OOXML/CLI integration | Windows .NET 10, no Office; during affected implementation work, exact focused `dotnet test` filter(s) documented by implementer | AC-01..07; EC-01a/b, EC-02a, EC-03a, EC-04a/b, EC-05a/b, EC-06a/b | Current source/fixture tests, source/PPTX SVG part checks, precise diagnostics, negative paths, output hashes |
| **VAL-02** | Tier 2 full repository | Windows .NET 10, run `./eng/validate.ps1` after integration stabilizes and at closure when stale | AC-01..07, DOC-01..03 | Full test/provenance/hygiene output, working revision and documentation-example validation |
| **VAL-03** | Tier 3 **two distinct native PPTX origin paths** | Existing authorized **interactive Windows** account with real desktop PowerPoint and LibreOffice Impress; invoke `./eng/test-m0020-pptx-tier3.ps1` (new script or equivalent named project entrypoint) once SVG output and fixtures are ready; no new host provisioning | AC-02, AC-08, EC-08a, EC-08b, REV-01 evidence inputs | Separate per-origin fixture/source/template/candidate hashes, versions, SVG part/vector proof, native slide renderings, visual fidelity/editability-as-one-picture and cleanup; missing runtime stays BLOCKED |
| **VAL-04** | Tier 4 **installed current tool consumer** | Windows .NET 10, isolated local package/tool directory, `./eng/test-m0020-tier4.ps1` (new script/equivalent), after public CLI/examples stabilize; native result is checked by VAL-03 target | AC-05, AC-09, DOC-01..03; EC-09a/b, EC-10a, EC-11a | Current `.nupkg` hash, isolated installation, `yadg init`, config/direct/selected commands, output hashes and native acceptance binding |
| **VAL-05** | Tier 5 **blocking human artifact review** | Human with PowerPoint + Impress on authorized Windows; `./eng/review-check.ps1 --milestone M0020` after current Tier 3/4 receipts and two candidates are generated | REV-01, AC-08, EC-08a, EC-08b | Genuine decision bound to exactly two PPTX artifacts, native-origin hashes, per-origin native renderings and receipt IDs; rejection returns to implementation |

VAL-01/02 are portable source-feedback/repository gates; they cannot stand in for VAL-03. VAL-03 and VAL-04 are independently resumable expensive acceptance stages. Run only missing/stale gates after relevant changes, with conservative invalidation on shared SVG/image packaging or fixture changes. Record revision, fixture origin/hashes, relevant test commands, current source/config/output/package hashes and real app versions; no stale M0019 artifact counts as M0020 proof. A truly absent native Office capability is a **BLOCKED gate**, not a pass, simulation or permission to install tools. Existing project-local Office automation isolation/cleanup and user-scope tool installation permissions apply. No new CI workflows/runners/secrets or persistent host provisioning is authorized.

## Blocking human review

- **ID:** `HR-M0020-01`, owner M0020, class `artifact-quality`, **blocking**, no waiver before completion.
- **Subjects:** `artifacts/review/evidence/M0020/deck-powerpoint-svg.pptx` and `artifacts/review/evidence/M0020/deck-impress-svg.pptx`; first based on a real PowerPoint-origin PPTX and reviewed/rendered in PowerPoint, second on a real Impress-origin PPTX and reviewed/rendered in Impress. Both must contain source-grid SVG diagrams exercising rectangles, connectors, nested structures and readable labels. No requirement to reapprove an unrelated Word document.
- **Evidence manifest:** `artifacts/review/evidence/M0020/review-manifest.json`; include candidate/fixture/source/config hashes, per-origin Tier 3 and Tier 4 receipt/package identity, native application versions, slide-by-slide native render hashes, SVG part presence/identity and editability. Human opens actual files and images and checks diagram legibility, fidelity, unambiguous layers, line/arrow display, aspect ratio, template geometry and no repair/overflow.
- **Approval:** record through established `.review/records/HR-M0020-01.md` and `./eng/review-check.ps1 --milestone M0020`. A reviewer/UTC timestamp, bound manifest and candidate hashes, explicit approved/rejected disposition are required; agents must not fabricate approval. Implement review-check and verifier compatibility for M0020 as part of scope.

## Documentation impact and completion

The implementation must author actual user documents in `docs/user/BOX-DIAGRAMS.md` and `docs/user/CONFIGURATION.md`, update `README.md`, affected authoring/presentation docs and doc indexes; tests must exercise samples rather than checking for presence alone. Update project-authority index `docs/SPECS.md` and milestone index `docs/MILESTONES.md` during implementation so the new spec/milestone are discoverable. This targeted authoring documentation work is **inside** M0020, not a deferred guide-sync pass. Broader unrelated documentation sync, release notes and release engineering remain deferred; record focused hints in `.guide-sync/pending/` only for a real remaining contradiction outside this milestone.

At closure reconcile every AC/DOC/REV obligation, every EC case, and all VAL receipts with `.execution/M0020.md`. Update this milestone's **Completion Evidence** section with concise results, final revision/validated fixture hashes and genuine review outcome. Keep the review gate pending if human sign-off is absent. Implementation must not copy guide-system documents into the repository or create an `EXECUTE-M0020.md` file.

### Completion Evidence

Implementation and documentation are complete. `./eng/validate.ps1` passed at closure with 142/142 integration tests, M0012/M0018 fixture verification and repository hygiene passing. VAL-03 passed independently for PowerPoint-origin→PowerPoint (16.0.20527.20196) and Impress-origin→Impress (26.8.2.2). VAL-04 passed with isolated Yadg 1.2.0 package SHA-256 `221DCDC62D1385A891156C32C13C8374612812F5B4E22CC3B13EB628CBD01A31`, including installed-tool native acceptance. Current review evidence is in `artifacts/review/evidence/M0020/review-manifest.json`; source revision at evidence capture is `2e1a5ccc9f72ec41566a5f51d0282e011e5f7b17`. Validated PowerPoint- and Impress-origin fixture hashes are `2DCBF9DBF91FBF1AEFFF9771CF9DAD1D3E66724DC39C409048B31620F93E52B7` and `033CAC3CB7956E5B511608CEC501DA3D21A050AE43FF13D1131358C8DFEA96A6`, respectively.

The final obligation reconciliation is recorded in `.execution/M0020.md`. All AC, DOC and EC implementation/validation obligations and VAL-01 through VAL-04 pass. Genuine human review `HR-M0020-01` remains pending (REV-01 and VAL-05); no reviewer approval has been fabricated. `./eng/review-check.ps1 --milestone M0020` correctly blocks completion until the reviewer inspects and records approval for both designated candidates.

## Implementation discretion and escalation

Implementation owns concrete classes/files/test structure, parser adapters, compliant SVG packaging and compatibility preview mechanics, validation scripts, documentation phrasing, run cadence and bounded work packages. Escalate to planning **before** changing the SVG-primary requirement, accepting raster-only output, adding native PowerPoint component shapes, changing public grammar/schema/layouts/YOLO failure semantics, introducing external tools or weakening the origin-matched native gates. For PPTX SVG interoperability, an auxiliary bitmap fallback within the same SVG-bearing picture object is permitted, but replacing the vector product with that fallback is not. A native runtime blocker remains explicitly BLOCKED.
