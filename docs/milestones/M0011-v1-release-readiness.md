# Milestone — M0011 V1.0 Release Readiness

## Correction status

This is the corrected M0011 release-readiness authority for the existing PR11 implementation branch.

It supersedes the earlier M0011 requirement to preserve MSBuild-generated Office interop assemblies.

Do not create a new milestone or parallel implementation branch solely for this correction.

## Release target

```text
YADG 1.0.0
```

M0011 produces a validated, human-approved release candidate and does not externally publish it.

## Execution profile

| Field | Value |
|---|---|
| Lifecycle state | ready / correction during implementation |
| Milestone type | release-readiness |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna |
| Repository role | product-tool |
| Release version | `1.0.0` |
| Distribution | one `Yadg` .NET tool |
| Platform | Windows x64 / .NET 10 |
| Word binding | built-in late-bound COM via `Word.Application` |
| Office interop binaries in package | forbidden |
| Tier 2 | `./eng/validate.ps1` |
| Tier 3 | `./eng/test-m0010-tier3.ps1` |
| Pack | `./eng/pack.ps1` |
| Tier 4 | `./eng/test-m0011-tier4.ps1` |
| Human review | narrow blocking `HR-M0011-01` |
| External publication | excluded |

## Goal

Correct the existing M0011 release candidate before V1 by:

1. replacing generated Office interop with late-bound Word COM;
2. simplifying build/package construction;
3. strengthening CLI/help release validation;
4. correcting README/changelog release surfaces;
5. regenerating exact release evidence;
6. narrowing the human review to actual human judgment.

## Required implementation correction

### Late-bound Word COM

Replace typed Office interop with:

```text
Type.GetTypeFromProgID("Word.Application")
-> Activator.CreateInstance
-> late-bound Word object model
```

Use `dynamic` or equivalent reflection-based late binding.

Remove:

```text
COMReference
ResolveComReference
tlbimp
Microsoft.Office.Interop.* references/packages
generated Office interop DLL dependencies
```

Do not switch to another wrapper technology.

Preserve the existing renderer's observable behavior, STA/timeout/lifecycle rules, real Word finalization semantics, and interactive-user restriction.

Do not introduce a helper process unless late binding cannot satisfy the existing real Word contract.

### Build/package correction

Restore ordinary .NET SDK restore/build/pack.

`eng/pack.ps1` follows `docs/engineering/RELEASE.md`.

The package must contain no Office interop wrapper assembly.

Existing PR11 package/evidence is invalid because it represents the superseded generated-interop package. Regenerate all M0011 package/evidence after the correction.

### CLI/help correction

Validate actual root and command-specific help.

Fix help descriptions/options so actual CLI help and README agree.

Evaluate/upgrade `System.CommandLine` to current stable 2.x as specified by release authority.

### README correction

Before completion:

- add a minimal complete onboarding example;
- include a concrete Mermaid producer config example;
- remove DMS wording;
- remove Visual Studio/`ResolveComReference`/`tlbimp` requirements from ordinary user troubleshooting;
- distinguish runtime prerequisites from building YADG itself;
- keep Word/LibreOffice behavior and limitations accurate;
- keep MIT wording scoped to YADG, not Microsoft/third-party software.

### CHANGELOG correction

The candidate is not yet released.

Use:

```text
## [Unreleased]
```

for the pending V1 entry.

Do not assign the actual `1.0.0` release date until a later publication/tag action.

## Existing behavior to preserve

Preserve:

- M0010 document composition/value behavior;
- real Word finalization;
- Word interactive-user restriction;
- real LibreOffice rendering;
- publish semantics;
- package ID `Yadg`;
- tool command `yadg`;
- version target `1.0.0`;
- Windows x64/.NET 10 support target;
- external-publication prohibition.

## Required authority

Read:

- `docs/engineering/RELEASE.md`
- `docs/ENGINEERING.md`
- `docs/SPECS.md`
- `docs/ARCHITECTURE.md`
- `docs/specs/WORD-RENDERER.md`
- `docs/specs/RENDERING.md`
- `docs/specs/PUBLISHING.md`
- `README.md`
- `AGENTS.md`
- `.review/pending/HR-M0011-01.md`

The corrected release authority supersedes older M0009/M0011 text that mandates generated Office interop for the release package.

Update directly contradicted project documentation so repository truth is coherent before completion.

## Acceptance criteria

### Word binding

- Word renderer has no compile-time `Microsoft.Office.Interop.*` / `Microsoft.Office.Core` dependency.
- Word renderer project has no Office `COMReference`.
- CLI project has no Office `COMReference`.
- no generated Office interop assembly is required to build or run the installed tool.
- activation uses version-independent `Word.Application`.
- late-bound implementation preserves existing Word renderer behavior/lifecycle/diagnostics.
- missing/unregistered Word fails actionably.
- real Word Tier-3/Tier-4 finalization succeeds.

### Build/package

- ordinary .NET SDK restore/build succeeds without Office type-library generation tooling.
- pack no longer requires Visual Studio `ResolveComReference`.
- exactly one `Yadg.1.0.0.nupkg` is produced.
- package contains no Office interop/generated wrapper.
- installed Word render succeeds despite wrapper absence.
- old package hash/evidence is not reused.

### CLI/help

- stable `System.CommandLine` 2.x is used unless a demonstrated incompatibility is escalated.
- `yadg --help` is correct.
- `yadg check --help` is correct.
- `yadg build --help` is correct.
- `yadg render --help` is correct.
- `yadg publish --help` is correct.
- version flag behavior is documented and tested consistently.
- README syntax/defaults agree with actual help.

### README/changelog

- README contains minimal complete onboarding.
- README contains concrete Mermaid config.
- README contains no DMS wording.
- README does not tell ordinary users they require interop-generation build tooling.
- README distinguishes renderer runtime prerequisites from build/release engineering.
- README license wording scopes MIT correctly.
- changelog uses `[Unreleased]`.

### Validation

Fresh corrected revision must pass:

```powershell
./eng/validate.ps1
./eng/test-m0010-tier3.ps1
./eng/pack.ps1
./eng/test-m0011-tier4.ps1
```

Local-only NuGet push-script validation also passes.

M0011 release evidence is regenerated for the corrected package/revision.

### Human review

`HR-M0011-01` asks only subjective release documentation/readiness questions.

Review tooling derives package hash, revision, runtime provenance, and automated results itself.

The human is not asked to transcribe SHA values, Word version, LibreOffice version, or test results.

`./eng/review-check.ps1 --milestone M0011` binds approval automatically to current release evidence.

## Human review

Review ID:

```text
HR-M0011-01
```

Human decides only:

- README acceptable for V1 users;
- pending V1 changelog/release description acceptable;
- corrected exact RC acceptable to mark release-ready.

Implementation does not fabricate approval.

## Completion flow

```text
apply correction overlay
-> amend existing PR11
-> replace generated interop with late-bound COM
-> simplify build/pack
-> improve CLI/help
-> correct README/changelog
-> fresh Tier 2
-> fresh M0010 Tier 3
-> fresh pack
-> fresh M0011 Tier 4
-> local publish-script validation
-> regenerate exact-hash release evidence
-> narrow HR-M0011-01
-> review-check
-> final authority/evidence/ledger reconciliation
```

## External publication

Still excluded.

Do not push NuGet externally, create a GitHub Release/tag/workflow, or otherwise publish `1.0.0`.

## Escalation boundary

Return to planning if:

- late-bound `Word.Application` cannot satisfy the existing real Word renderer contract;
- late binding requires changing document finalization semantics;
- another COM binding strategy/helper process appears necessary;
- stable `System.CommandLine` 2.x creates a material CLI incompatibility needing product policy;
- ordinary .NET tool packaging cannot package/install the corrected tool;
- M0010 Tier-3 behavior regresses in a way requiring a semantic change.

Do not fall back to generated Office interop inside M0011.
