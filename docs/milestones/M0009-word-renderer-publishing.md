# Milestone — M0009 Microsoft Word Renderer and Publishing

## Execution Profile

| Field | Value |
|---|---|
| Lifecycle state | ready |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna |
| Baseline executor readiness | confirmed |
| Decision preservation | confirmed |
| Execution tractability | confirmed |
| Scope size | large |
| Implementation autonomy | high |
| Documentation sync | deferred |
| Repository validation | `./eng/validate.ps1` on Windows |
| Integration validation | real Microsoft Word + real filesystem publishing |
| Validation locus | Windows; real Word target requires interactive logged-on user session |
| Consumer/release validation | deferred to M0010 |
| Human review | blocking `HR-M0009-01` |

## Goal

Complete the document lifecycle needed before V1 release-readiness work:

1. add Microsoft Word as a real renderer/finalizer for authored DOCX;
2. add an explicit `publish` command that copies finalized DOCX from workspace state to a caller/configured delivery filesystem path.

M0009 does not perform release/NuGet packaging; that follows in M0010.

## Target State

```text
Markdown / YADG.md / templates
        |
        | yadg build
        v
YadgPreWords/*.docx
        |
        | yadg render --renderer word
        v
YadgWords/*.docx
        |
        | yadg publish --publish-path <path>
        v
delivery path/*.docx
```

Microsoft Word is the V1 fidelity target.

LibreOffice remains supported/default and retains its existing PDF side output.

Publication is renderer-independent and DOCX-only in M0009.

## Scope

### Microsoft Word renderer

- renderer ID `word`;
- real desktop Microsoft Word COM automation;
- Windows authoritative build/runtime locus;
- isolated Word-specific dependency boundary;
- no requirement to preserve Linux full-solution build;
- new owned Word application instance per render invocation;
- field/index/pagination refresh;
- save finalized DOCX to `YadgWords`;
- runtime/provenance diagnostics;
- bounded failure/cleanup behavior;
- real Word Tier-3 integration;
- blocking artifact-quality review.

### Publishing

- `yadg publish`;
- optional `YADG.md publish.path`;
- `--publish-path` CLI override;
- workspace-root relative-path resolution;
- all top-level finalized `YadgWords/*.docx`;
- same-name replacement/unrelated-file preservation;
- staging/temp copy hygiene;
- no implicit build/render;
- no PDF publication.

## Non-goals

M0009 does not implement:

- NuGet/.NET tool packaging or publishing;
- GitHub workflows/releases;
- DMS integration;
- ZIP/release packages/manifests/signing;
- PDF publication;
- Microsoft Word PDF export;
- server-side/non-interactive Word automation support;
- remote Word automation;
- automatic Office install/licensing;
- generic renderer-equivalence guarantees;
- arbitrary publication filtering/renaming;
- Linux full-product support guarantee.

## Resolved Decisions and Constraints

### Renderer compatibility

Supported IDs:

```text
libreoffice
word
```

Default remains `libreoffice`.

`--renderer-path` is valid only for LibreOffice.

Word outputs finalized DOCX only.

LibreOffice preserves existing DOCX + PDF behavior.

### Microsoft Word automation locus

Authoritative runtime:

```text
Windows
+ installed desktop Microsoft Word
+ interactive logged-on user
+ normal user profile
+ Office activation/first-run complete
```

The milestone does not claim support for service/SYSTEM/server-side automation.

### Build/platform contract

M0009 allows the full product/CLI to become Windows-specific.

Office-independent components must keep Word automation dependencies isolated.

Linux build preservation is not required for completion.

The exact managed COM binding strategy is implementation-owned. PIA/interop metadata from the Windows environment, late-bound/dynamic COM, or another valid Windows COM technique may be used.

Do not establish an unverified NuGet Office interop package as project authority merely for portability.

If no workable modern .NET binding can be produced on the required real Windows/Office locus without changing these architectural boundaries, escalate to planning.

### Word ownership/finalization

YADG creates/owns a Word automation instance and does not attach to an existing user instance.

Documents are processed serially.

Authored inputs are never edited in place.

Observable correctness, not a prescribed sequence of COM calls, defines finalization.

Required current results include YADG `SEQ`/`REF`/section-number references and representative template-owned TOC/list structures.

Word saves the final DOCX.

### Publishing destination

CLI:

```text
yadg publish [--workspace <path>] [--publish-path <path>]
```

Precedence:

```text
--publish-path
    >
YADG.md publish.path
    >
error
```

Relative CLI/config paths resolve against workspace root.

Destination may be outside workspace or a safe workspace-local directory such as `./Published`.

Destination must not be a reserved YADG artifact/source directory or descendant.

### Published artifacts

M0009 publishes:

```text
YadgWords/*.docx
```

top-level only, all documents, basenames unchanged.

It does not publish `YadgPdfs`.

### Automation separation

`publish` contains no Word automation dependency and is suitable for ordinary filesystem automation/CI.

`publish` never runs `build` or `render`.

## Baseline Executor Readiness

Planning has settled:

- supported renderer ID/default;
- Windows/interactive Word specialization;
- allowed build-platform change;
- interop binding decision boundary;
- Word output artifact set;
- Word lifecycle/finalization responsibilities;
- publication CLI/config precedence;
- destination/path semantics;
- published artifact set;
- failure/commit semantics;
- validation targets/loci;
- human-review gate;
- M0010 boundary.

Implementation owns concrete projects/types, COM binding/reflection/PIA mechanics, exact Word automation call sequence, process/watchdog mechanics, diagnostic codes, publish file-copy helpers, test fixture construction, and execution decomposition.

## Execution Tractability

Create and maintain:

```text
.execution/M0009-word-renderer-publishing.md
```

Use bounded work packages and map every acceptance criterion to concrete evidence.

## Required Authority

Read before implementation:

- `docs/SPECS.md`
- `docs/specs/RENDERING.md`
- `docs/specs/WORD-RENDERER.md`
- `docs/specs/PUBLISHING.md`
- `docs/specs/WORKSPACE-VALUES.md`
- `docs/specs/WORD-REFERENCES.md`
- `docs/specs/STRUCTURED-CONTENT.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`
- `.review/pending/HR-M0009-01.md`

Do not require the implementation agent to read the external guide repository or planning conversation.

## Acceptance Criteria

### Renderer selection/platform

- `render --renderer word` is recognized.
- existing default remains LibreOffice.
- `--renderer-path` with Word fails clearly.
- Word-specific runtime/build dependencies do not leak into authoring-core/Open XML semantic layers.
- authoritative full solution builds on the declared Windows development locus.
- no completion claim depends on Linux building successfully.

### Word runtime preflight

- missing/unregistered Word produces actionable failure before normal output commit.
- Word version is obtained/reported for real-runtime evidence.
- renderer uses an owned application instance rather than attaching to unrelated interactive Word.
- automation is non-visible/suppresses ordinary prompts where available.
- macro execution is suppressed where practical for supported trusted DOCX handling.
- bounded failure handling prevents an indefinitely hanging CLI.

### Word finalization

- every top-level `YadgPreWords/*.docx` gets a corresponding Word-finalized `YadgWords/*.docx`.
- authored inputs remain untouched.
- representative `SEQ` displayed values are current.
- representative `REF` displayed values are current.
- numbered section references are current.
- template-owned TOC is updated where present in the integration fixture.
- template-owned list-of-figures/list-of-tables is updated where present/appropriate.
- pagination-dependent results used by fields/indexes are current.
- output opens in Word without repair.
- no Word PDF output is created/required.
- existing unrelated `YadgWords` files are preserved.
- failed/incomplete temp outputs are not presented as success.
- documents/application are closed/quitted on success and best-effort failure.
- real test evidence demonstrates no owned Word process leak under normal success and a controlled failure scenario where practical.

### LibreOffice compatibility

- existing LibreOffice renderer behavior/tests remain materially intact.
- LibreOffice remains default.
- existing LibreOffice PDF behavior is not removed merely because PDF is outside M0009 publishing.

### Workspace publish configuration

- schema-v1 `YADG.md` accepts optional `publish.path`.
- existing workspaces without `publish` remain valid.
- `publish.path` is a non-empty string.
- unknown publish keys fail.
- `check` validates schema but does not require destination existence/writability.

### Publish command/path precedence

- CLI `--publish-path` overrides workspace default.
- configured default is used when CLI path absent.
- missing both causes `publish` failure.
- relative CLI/config paths resolve against workspace root.
- external absolute filesystem destinations work.
- safe workspace-local destination works.
- reserved YADG directories/descendants are rejected.

### Publish artifact behavior

- requires at least one top-level `YadgWords/*.docx`.
- publishes all such DOCX without recursion.
- preserves basenames.
- creates destination when needed.
- same-name destination files are replaced.
- unrelated destination files are preserved.
- templates, PreWords, PDFs, producer temp assets are not published.
- publish never invokes build/render and never modifies DOCX contents.

### Publish failure hygiene

- preflight failures do not change destination files.
- each copy uses destination-local temporary staging before final filename commit.
- incomplete temp files are cleaned best-effort.
- cross-document rollback after later sibling failure is not required.

### Documentation/hygiene

- README/direct public usage documentation covers Word renderer prerequisites/invocation and publish path precedence.
- docs do not claim DMS/PDF publication/server-side Word support.
- fixtures are synthetic/non-confidential.
- execution ledger maps every criterion to evidence.

### Human review

- implementation prepares required review evidence for `HR-M0009-01`.
- existing review tooling is extended to recognize M0009 without weakening M0005 historical checks.
- `./eng/review-check.ps1 --milestone M0009` fails pending/negative review.
- approval cannot be fabricated by implementation.
- no waiver path.
- milestone remains awaiting review until human approval exists.

## Validation

### Tier 0 — edit sanity

On Windows:

```powershell
dotnet build Yadg.slnx --configuration Release
```

or the current repository-equivalent command if solution targeting changes.

Evidence must include the Word-enabled product.

### Tier 1 — focused validation

Cover:

- renderer option routing;
- Word preflight/failure surfaces with substitutes where real Word is unnecessary;
- LibreOffice regression;
- publish front-matter parsing;
- path precedence/resolution;
- reserved path rejection;
- filesystem copy/overwrite/preservation/failure hygiene.

### Tier 2 — repository validation

On the authoritative Windows locus:

```powershell
./eng/validate.ps1
```

M0009 may change this script's platform prerequisites to Windows/Word-development capability if compilation requires it.

Tier 2 does not substitute for real Word Tier 3.

### Tier 3A — real Microsoft Word integration

Required capability:

```text
interactive Windows user session
desktop Microsoft Word installed/activated
repository .NET SDK/build prerequisites
```

Use a synthetic workspace/template exercising:

- numbered Markdown section/reference;
- figure caption/SEQ/REF;
- table caption/SEQ/REF where practical;
- template-owned TOC;
- template-owned list of figures/tables where practical;
- ordinary content preserving template formatting.

Run real:

```powershell
yadg build --workspace <workspace>
yadg render --workspace <workspace> --renderer word
```

The exact launcher may be `dotnet run` before packaging.

Evidence inspects the real Word-saved DOCX and proves current displayed field/index values and preserved structures.

Record OS, Word version, source/final artifact identity/hash, and command result.

LibreOffice, mocks, or direct OOXML rewriting are not replacement evidence.

### Tier 3B — real publishing integration

Using a synthetic workspace with multiple finalized DOCX:

- run with configured workspace-local `publish.path`;
- run with CLI override to another temporary path;
- inspect delivered bytes/names;
- verify PDF/intermediate/template exclusion;
- verify replacement/preservation;
- verify failure cases.

Word installation is not conceptually required for publish semantics, although the M0009 product build may be Windows-bound.

### Tier 4 — consumer/release

Deferred to M0010.

### Tier 5 — blocking human review

Review:

```text
HR-M0009-01
```

Implementation prepares representative real Word artifact/evidence.

The human reviewer follows `.review/pending/HR-M0009-01.md`.

Completion requires:

```powershell
./eng/review-check.ps1 --milestone M0009
```

to pass with an approved record.

## Constrained Runtime

The implementation agent may not have real Microsoft Word available.

If Windows/Word capability is absent:

- implement and run all validation that does not require it;
- do not substitute LibreOffice/mock evidence for Tier 3A;
- leave the ledger explicitly blocked on the real Word target rather than claiming milestone completion.

If real Word is available, Tier 3A and review evidence preparation are mandatory.

Human approval itself must be performed by the user/reviewer, not the implementation agent.

## Research

Planning revalidated current Microsoft guidance sufficiently to establish the project specialization:

- Office/Word automation is COM-based;
- Microsoft documents Office PIAs for managed Office automation/development;
- modern .NET documentation distinguishes legacy PIA/type-library guidance from newer COM mechanisms;
- unattended/non-interactive Office automation is not a supported project execution target.

No durable `docs/research/` artifact is created because implementation-affecting conclusions are promoted into project authority and research is inactive in the repository profile.

Implementation need not reconstruct this research.

## Documentation Impact

Planning adds/updates:

- `docs/specs/WORD-RENDERER.md`
- `docs/specs/PUBLISHING.md`
- `docs/specs/RENDERING.md`
- `docs/specs/WORKSPACE-VALUES.md`
- `docs/SPECS.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`
- `docs/MILESTONES.md`
- `.review/pending/HR-M0009-01.md`
- this milestone

M0008 is reconciled to complete.

Implementation updates direct README/usage docs and extends review tooling for M0009.

Broad V1 documentation review is deferred to M0010.

No GitHub workflow work is included.

No `.guide-sync/pending/` hint is required.

## Human Review

Applicability: `blocking`

Review class: `artifact-quality`

Review ID: `HR-M0009-01`

Owning milestone: `M0009`

Waiver: forbidden.

The request is `.review/pending/HR-M0009-01.md`.

Implementation extends milestone-scoped review tooling and prepares evidence; only the human reviewer records the decision.

## Completion Expectations

Implementation owns:

```text
read milestone/authority
-> bounded execution decomposition
-> create/update execution ledger
-> implement Word renderer
-> implement publish
-> Tier 0-2
-> real Word Tier 3A
-> real publish Tier 3B
-> prepare human review evidence
-> obtain human review decision externally
-> review-check
-> fresh authority reread
-> milestone <-> ledger <-> repository/evidence reconciliation
-> completion audit
```

Passing automated tests alone is insufficient because `HR-M0009-01` is blocking.

## Baseline-Executability Audit

Planning confirms:

- M0008 completion evidence exists;
- applicable profiles remain repository-wide `base` + `artifact-first-runtime`;
- repository role remains `product-tool`;
- M0009 is ordinary implementation, not release-readiness;
- Windows/Word specialization is explicit;
- loss of Linux full-build guarantee is explicit;
- interop binding style is an implementation mechanic, not unresolved product policy;
- renderer selection/default/output semantics are fixed;
- publication CLI/config/path/artifact semantics are fixed;
- PDF/DMS/release boundaries are fixed;
- validation targets/loci/fallback semantics are fixed;
- human review subject/evidence/gate is fixed;
- M0010 boundary is fixed.

Remaining choices are implementation mechanics.

M0009 is ready for GPT-5.6 Luna.

## Escalation Boundary

Return to planning if implementation requires a material change to:

- Microsoft Word runtime or interactive-user locus;
- Word renderer ownership/finalization semantics;
- renderer ID/default;
- Word PDF scope;
- authoring-core/renderer dependency boundary;
- product platform support commitment;
- use of a new mandatory third-party interop/runtime dependency;
- publish CLI/config precedence;
- publication source/destination/artifact semantics;
- implicit build/render behavior;
- DMS/remote publication;
- validation target/locus;
- blocking human-review policy.

Failure to find any workable modern .NET binding to real Word on the declared Windows/Office locus is an escalation, not permission to replace real Word with another renderer.
