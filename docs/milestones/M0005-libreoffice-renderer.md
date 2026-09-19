# Milestone — M0005 LibreOffice Renderer and Finalization

## Execution Profile

| Field | Value |
|---|---|
| Lifecycle state | ready |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna |
| Baseline executor readiness | confirmed |
| Decision preservation | confirmed |
| Execution tractability | confirmed |
| Scope size | medium-large |
| Implementation autonomy | high |
| Documentation sync | deferred |
| Focused validation | renderer process/API, refresh/save/export, failure semantics, and review-workflow tests |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real installed LibreOffice + real M0004-authored DOCX + finalized DOCX/PDF inspection |
| Validation locus/platform | local/CI only where an actual LibreOffice runtime is installed and API-controllable |
| Consumer/release validation | not applicable; packaging unresolved |
| Human review | blocking: `HR-M0005-01` artifact-quality review |

## Goal

Implement the first real YADG renderer/finalizer using LibreOffice.

A user can take Office-independent `YadgPreWords/*.docx`, run `yadg render`, and receive finalized DOCX plus PDF artifacts whose M0004 fields/indexes have been evaluated by a real LibreOffice runtime.

M0005 also introduces the repository-local human-review substrate and gates milestone completion on review of a representative rendered artifact.

## Target State

When M0005 is complete:

- `yadg render` is a real workspace command;
- the default/only M0005 renderer ID is `libreoffice`;
- renderer executable path can be supplied explicitly or discovered from `PATH`;
- top-level PreWords are rendered without implicit rebuilding;
- each invocation uses an isolated temporary LibreOffice profile/session;
- M0004 `SEQ`/`REF` values and existing document indexes are refreshed using LibreOffice's programmatic API;
- finalized DOCX files are written to `YadgWords/`;
- corresponding PDFs are written to `YadgPdfs/` from the same refreshed document state;
- runtime version/platform provenance is observable and recorded in validation/review evidence;
- Office-independent Tier 2 remains runnable without LibreOffice;
- real LibreOffice Tier 3 proves field/index refresh and DOCX/PDF output;
- blocking review `HR-M0005-01` is approved by a human before milestone completion.

## Scope

M0005 covers:

- renderer abstraction sufficient for future additional engines;
- LibreOffice concrete renderer;
- `render --workspace`, `--renderer`, and `--renderer-path` CLI behavior;
- PreWord input discovery;
- `YadgWords`/`YadgPdfs` artifact conventions;
- isolated LibreOffice launch/profile/API session lifecycle;
- runtime version reporting;
- text-field refresh;
- Writer document-index update;
- finalized DOCX save;
- PDF export;
- failure/cleanup semantics;
- real-runtime integration validation;
- milestone-scoped review request/record/check infrastructure;
- synthetic human-review evidence generation;
- documentation needed to use/install/provide the renderer runtime.

## Non-goals

M0005 does **not** implement:

- Microsoft Word renderer/Office Interop;
- claims of LibreOffice/Microsoft Word visual equivalence;
- `publish`;
- DMS integration;
- package/installer distribution;
- automatic LibreOffice installation/update;
- remote renderer services;
- PDF/A/signing/encryption/watermark/publication policy;
- configurable PDF filter settings;
- source-authoring semantic changes;
- prepared-table row population;
- scalar `value` resolution;
- broad documentation synchronization.

## Decisions and Constraints

- Applicable guide profiles remain repository-wide `base` and `artifact-first-runtime`; no profile conflict exists.
- YADG remains a product/tool.
- The durable renderer boundary is engine-neutral; LibreOffice is the first implementation.
- `render` does not invoke `build` implicitly.
- Input/output conventions and CLI semantics are fixed by `docs/specs/RENDERING.md`.
- Finalized DOCX and PDF for one input come from the same refreshed loaded LibreOffice document state.
- LibreOffice is controlled programmatically; plain DOCX command-line round-tripping alone is not accepted as finalization evidence.
- Every invocation uses a dedicated temporary user profile and does not attach to an existing user LibreOffice instance.
- Runtime connection is local-only and requires no credentials/network service.
- No hard-coded minimum version is product authority; exact runtime version/platform is provenance and must pass real capability validation.
- Tier 2 remains renderer-independent/portable.
- Tier 3 LibreOffice validation is authoritative for renderer behavior and cannot be replaced by mocks or OOXML-only assertions.
- Human review is blocking and has no waiver path in M0005.

## Planning Evidence / Research Boundary

Planning used official LibreOffice command/API documentation and one local synthetic interoperability probe.

Direct observations:

- LibreOffice supports `--headless`, API acceptors, and an isolated profile via `-env:UserInstallation=...`.
- Writer text-field collections expose refresh capability; Writer document indexes expose update/refresh capability.
- Writer PDF export is available through `writer_pdf_Export`.
- On LibreOffice 25.2.3.2 (Linux, planning environment), a synthetic Word-style `SEQ` + bookmark + `REF` DOCX exported to PDF with evaluated value `1`, while a plain DOCX-to-DOCX command-line conversion preserved cached field result `0` in the saved OOXML.

Planning conclusion promoted into project authority: M0005 must explicitly refresh/update fields/indexes through the programmatic renderer session and validate the finalized DOCX; plain conversion is not sufficient.

No durable `docs/research/` artifact is created because the current repository profile does not activate that documentation layer and the operative conclusions are fully promoted into project authority. The implementation must record the exact real runtime used for M0005 evidence.

## Baseline Executor Readiness

Planning has settled:

- renderer architecture and first engine;
- public CLI selection/path semantics;
- input/output artifact locations;
- build/render separation;
- field/index refresh obligation;
- same-session DOCX/PDF requirement;
- process/profile isolation;
- runtime discovery/provenance and fallback behavior;
- validation target/locus;
- human review ID/class/evidence/decision policy;
- review-check command and completion behavior.

Implementation owns:

- concrete project/type/file structure;
- UNO/.NET bridge or sidecar mechanics consistent with architecture;
- local-only UNO transport and handshake details;
- exact refresh/update call ordering needed to satisfy evidence;
- temporary staging and atomic replacement mechanics;
- executable PATH search details;
- bounded timeout values;
- diagnostic-code allocation;
- test decomposition/fixture helpers;
- review-command implementation details consistent with project authority.

## Execution Tractability

At implementation start create and maintain:

```text
.execution/M0005-libreoffice-renderer.md
```

The ledger maps all acceptance/validation/review obligations to bounded work packages and evidence.

Because Tier 3 requires an external installed runtime and Tier 5 requires a human decision, the executor must keep those gates explicit rather than conflating code completion with milestone completion.

## Required Authority

Read before implementation:

- `docs/SPECS.md`
- `docs/specs/STRUCTURED-CONTENT.md`
- `docs/specs/WORD-REFERENCES.md`
- `docs/specs/RENDERING.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`
- `.review/pending/HR-M0005-01.md` when producing/reconciling the required human-review evidence

Do not require the external guide repository or planning conversation.

## Acceptance Criteria

### CLI and artifact lifecycle

- `yadg render [--workspace <path>] [--renderer libreoffice] [--renderer-path <path>]` implements the specified contract.
- Omitted renderer defaults to `libreoffice`; unsupported renderer IDs fail clearly.
- Explicit renderer path takes precedence; otherwise LibreOffice is discovered from `PATH`.
- Render requires at least one top-level `YadgPreWords/*.docx` and does not recurse.
- Render does not modify PreWords and does not implicitly run build.
- Successful render produces same-named `YadgWords/*.docx` and same-stem `YadgPdfs/*.pdf`.
- Existing corresponding outputs may be replaced; unrelated output files are preserved.

### Runtime isolation and lifecycle

- Each render invocation uses a unique temporary LibreOffice user profile.
- The renderer starts/owns its automation process/session and does not attach to an unrelated interactive LibreOffice process.
- The automation channel is local-only.
- Runtime version is obtained and exposed in diagnostics/log/evidence.
- Renderer process/profile resources are cleaned on success and best-effort on failure.
- Missing/unstartable/incompatible runtime fails with actionable diagnostics rather than hanging.

### Field/index finalization

Using real M0004-authored DOCX:

- figure/table `SEQ` captions evaluate to current non-placeholder numbers;
- `REF` references evaluate to their target numbers;
- numbered section references evaluate correctly for the synthetic scenario;
- existing template-owned Writer document indexes are updated where present;
- no M0004 semantic target is silently rewritten to avoid LibreOffice incompatibility;
- field/index structures needed for subsequent editing remain present in finalized DOCX where the validated LibreOffice DOCX export supports them.

### Finalized DOCX and PDF

- Finalized DOCX contains current displayed values for the required synthetic reference scenario rather than M0004 placeholder/cached values.
- PDF displays the same resolved caption/reference values as the finalized DOCX.
- Finalized DOCX and PDF are produced from the same refreshed loaded-document state.
- Unrelated template content, generated tables/figures/styles, and ordinary document structure remain materially intact.
- Output files are non-empty, openable by the target runtime, and not temporary/incomplete artifacts presented as success.

### Failure semantics

- Workspace/input/renderer availability preflight occurs before normal outputs are modified.
- Failure to open, refresh, update an index, save DOCX, or export PDF returns non-zero with actionable context.
- Failure after sibling outputs have completed does not require transactional rollback, but the failed artifact is not reported as successful.
- Bounded waits prevent indefinite process hangs.

### Portability boundary

- `./eng/validate.ps1` remains runnable without LibreOffice.
- Renderer-specific dependencies do not leak into core semantic/OOXML authoring components.
- A fake renderer does not satisfy Tier 3.
- No Microsoft Office dependency is introduced.

### Review substrate

- `.review/pending/HR-M0005-01.md` is honored as the canonical blocking review request.
- Repository-local review commands support at least listing/showing M0005 review state, recording a human decision, and milestone-scoped checking.
- `./eng/review-check.ps1 --milestone M0005` fails before approval and passes only after an acceptable human record exists.
- An implementation agent cannot create/claim human approval on its own.
- Review records capture milestone/review ID, decision, reviewer identity, repository revision, LibreOffice version/platform, and evidence identity/hash where practical.
- Completed M0005 review is historical and is not designed to become stale after later commits.

### Documentation/hygiene

- Public documentation describes how `render` obtains LibreOffice and where outputs are written.
- The repository ends M0005 without contradictory Microsoft-Word-first renderer claims.
- Synthetic fixtures/evidence contain no confidential material.
- `.execution/M0005-libreoffice-renderer.md` maps every acceptance/review obligation to evidence.

## Validation

### Tier 0 — edit sanity

Target: changed .NET/repository review metadata and docs.

Locus: ordinary local/CI .NET environment.

Use the live solution and current repository build conventions.

### Tier 1 — focused validation

Cover renderer command parsing, runtime discovery/path precedence, process/profile isolation mechanics, failure diagnostics, output staging, and review-check mechanics.

Focused tests may use controllable process abstractions where appropriate, but they are not renderer-integration evidence.

Exact commands are executor-owned and recorded in the ledger.

### Tier 2 — repository validation

```powershell
./eng/validate.ps1
```

Must pass without requiring LibreOffice.

### Tier 3 — LibreOffice integration

Authoritative target: an actual installed LibreOffice runtime controlled headlessly/programmatically.

Required evidence uses a synthetic M0004 workspace/artifact with at least:

- one numbered Markdown section reference;
- numbered figure caption + `REF`;
- numbered table caption + `REF`;
- one existing template-owned document index (TOC required; figure/table list when practical for the fixture);
- enough ordinary M0003 content to detect destructive round-trip behavior.

Representative flow:

```powershell
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- build --workspace <workspace>
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- render --workspace <workspace> --renderer libreoffice [--renderer-path <path>]
```

Evidence includes:

- exact LibreOffice version/platform/executable;
- exit status and diagnostics;
- input/output artifact hashes or equivalent identity;
- structural inspection of finalized DOCX field/bookmark/index preservation;
- inspection proving cached/displayed required values are no longer placeholders;
- PDF text/content checks for corresponding evaluated numbers;
- successful reopening of finalized DOCX by LibreOffice;
- isolation evidence showing a dedicated user profile was used.

If the real runtime is unavailable at the executor's locus, Tier 0-2 may pass but the milestone remains `BLOCKED`; substitute conversion/mocking cannot close Tier 3.

### Tier 4 — consumer/release validation

Not applicable; packaging remains unresolved.

### Tier 5 — blocking human review

Review ID: `HR-M0005-01`

Class: `artifact-quality`

Owning milestone: M0005

Subject: representative LibreOffice-finalized YADG DOCX/PDF.

Evidence directory:

```text
artifacts/review/evidence/M0005/
```

Required evidence is defined in `.review/pending/HR-M0005-01.md` and must include the finalized DOCX, PDF, review manifest/provenance, and their identities/hashes.

Reviewer role: project maintainer/user capable of judging whether the rendered template artifact is visually and semantically usable.

Acceptable completion decision: `approved` only.

`changes-requested` returns implementation to milestone-scoped correction/evidence regeneration. `rejected` blocks completion. Waiver is not permitted.

Canonical gate command after the human decision:

```powershell
./eng/review-check.ps1 --milestone M0005
```

The implementation agent must terminate `AWAITING HUMAN REVIEW` when all agent-resolvable work and Tier 0-3 validation are complete but this review remains pending.

## Constrained Runtime

Portable validation must remain bounded and independent of an installed renderer.

Real LibreOffice integration may run only where the runtime exists. Use one isolated profile/session per invocation and bounded process waits; do not evade execution-harness limits with detached/background processes.

The human review is intentionally outside automated completion and must not be fabricated by the executor.

## Documentation Impact

This planning package also restores the M0004 authority files that were supplied to M4 implementation but not applied to `main`.

Planning updates/adds:

- `docs/SPECS.md`
- `docs/specs/WORD-REFERENCES.md` (M4 authority restoration)
- `docs/specs/RENDERING.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`
- `docs/MILESTONES.md`
- `docs/milestones/M0004-references-word-structures.md` (historical authority restoration)
- this milestone
- `.review/pending/HR-M0005-01.md`

Implementation must update `README.md` and any directly contradictory renderer documentation.

No `.guide-sync/pending/` hint is required.

## Human Review

Applicability: **blocking**.

Review class: `artifact-quality`.

Canonical ID: `HR-M0005-01`.

Request: `.review/pending/HR-M0005-01.md`.

Record: `.review/records/HR-M0005-01.md` after a human decision.

No waiver is permitted for M0005.

The review gates M0005 only and does not become perpetual approval of later renderer changes.

## Completion Expectations

Implementation owns:

```text
execution decomposition
-> persistent ledger
-> implementation + Tier 0-3 validation
-> review evidence generation
-> AWAITING HUMAN REVIEW if pending
-> human decision
-> milestone-scoped review-check
-> fresh milestone/authority reread
-> milestone <-> ledger <-> repository/evidence reconciliation
-> completion audit
```

M0005 cannot be `COMPLETE` merely because automated validation succeeds.

## Baseline-Executability Audit

Planning confirms:

- M4 project authority is restored for disconnected execution;
- renderer architecture/CLI/artifact semantics are fixed;
- LibreOffice runtime target and isolation are fixed;
- field/index refresh obligations are fixed;
- DOCX/PDF same-session semantics are fixed;
- runtime availability/version provenance and fallback behavior are explicit;
- validation depth and real target/locus are explicit;
- blocking human review has an ID, class, subject, evidence, reviewer role, decisions, no-waiver policy, paths, and exact review-check command;
- no research conclusion remains only in planning context;
- remaining decisions are local implementation mechanics;
- implementation can resume from repository-local ledger/review state.

M0005 is ready for GPT-5.6 Luna, subject to availability of the declared real LibreOffice integration target for Tier 3.

## Escalation Boundary

Return to planning if implementation requires a material change to:

- renderer CLI or artifact directories;
- build/render phase separation;
- LibreOffice-versus-engine-neutral architecture;
- field/index semantic expectations;
- same-session DOCX/PDF requirement;
- runtime isolation/security boundary;
- version/platform support policy;
- validation target/locus or fallback semantics;
- human review subject/class/acceptance/waiver policy;
- M0004 semantic contracts to work around renderer incompatibility.

Implementation owns concrete API/bridge/process/test mechanics that preserve this contract.
