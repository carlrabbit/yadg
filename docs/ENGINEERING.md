# Engineering

## Scope

Repository-wide engineering and validation policy for YADG.

## Product boundary

YADG is C#/.NET.

Open XML authoring remains separate from real document renderers.

M0009 introduces a Windows/desktop-Microsoft-Word specialization and explicit publishing.

## Platform policy after M0009

The authoritative full-product build and validation locus is Windows.

A complete Linux build is no longer a release/readiness requirement.

Authoring components should remain free of Word COM dependencies, but the CLI/full solution may become Windows-targeted if required by the selected interop mechanism.

Do not add complexity solely to preserve Linux compilation if that conflicts with a robust Word renderer.

## Microsoft Word development/build capability

The authoritative Word-capable development/build machine has:

- Windows;
- repository-supported .NET SDK;
- desktop Microsoft Word installed;
- the Office/COM registration/interoperability metadata required by the chosen implementation mechanism;
- an interactive user profile with Office activation/first-run complete.

The exact interop binding mechanism is implementation-owned.

Do not treat a third-party/unverified NuGet interop package as authoritative merely to preserve portability.

## Word automation execution policy

Authoritative real Word execution occurs in an interactive logged-on Windows user session.

Do not claim support for automation from a Windows service, SYSTEM Task Scheduler job, ASP/ASP.NET host, DCOM server, or other non-interactive/server-side context.

Word tests/runs must avoid concurrent YADG-owned Word automation instances.

Tests must clean up owned documents/application instances and verify, where practical, that no owned Word process remains.

## Canonical engineering interface

The canonical repository validation command remains:

```powershell
./eng/validate.ps1
```

After M0009 its authoritative locus is Windows.

Implementation may keep non-Windows subsets working, but they are not M0009 completion evidence for the Word-enabled product.

## Validation tiers

### Tier 0 — edit sanity

Windows build/static/config checks for changed areas.

### Tier 1 — focused

Renderer orchestration/failure handling and publish path/copy semantics.

Mocks/fakes may cover paths that do not claim real Word behavior.

### Tier 2 — repository

```powershell
./eng/validate.ps1
```

Runs on the authoritative Windows build locus.

Real Word invocation may remain separate in Tier 3, but the Word-enabled product must compile.

### Tier 3 — integration

M0009 has two real targets:

1. real Microsoft Word on an interactive Windows user session;
2. real filesystem publication against finalized DOCX inputs.

Mocks/fakes/OOXML-only inspection are not equivalent evidence for Word automation.

### Tier 4 — consumer/release

Deferred to the V1 release-readiness milestone.

### Tier 5 — human review

M0009 owns blocking artifact-quality review `HR-M0009-01`.

## Word Tier-3 evidence

A real Word integration run records at least:

- OS/platform;
- observable Word version;
- repository revision;
- authored DOCX identity/hash where practical;
- finalized DOCX identity/hash;
- command/result.

The fixture exercises YADG-owned structures whose correctness depends on finalization: sequence numbering, REF fields, numbered section references, and template-owned indexes where feasible.

Structural post-save inspection complements but does not replace opening/refreshing/saving through real Word.

## Publishing validation

Publishing tests use real temporary filesystem destinations.

Cover:

- configured default path;
- CLI override precedence;
- workspace-relative resolution;
- external absolute destination;
- reserved-directory rejection;
- all top-level `YadgWords/*.docx`;
- no PDF/intermediate/template copying;
- replacement of same-name destination;
- preservation of unrelated files;
- no implicit render/build;
- failure without effective destination;
- failure without finalized DOCX;
- temp/incomplete-output hygiene.

## M0009 human review

Canonical review ID:

```text
HR-M0009-01
```

Review class:

```text
artifact-quality
```

Owning milestone:

```text
M0009
```

Implementation extends existing milestone-scoped review tooling so:

```powershell
./eng/review-check.ps1 --milestone M0009
```

fails until an acceptable human record exists.

No implementation agent may fabricate approval.

The review record identifies at least:

- milestone/review ID;
- decision/status;
- human reviewer;
- repository revision;
- Microsoft Word version;
- reviewed finalized DOCX hash/evidence identity.

Waiver is forbidden.

## Human review subject

The reviewer opens a representative M0009 finalized DOCX in Microsoft Word and checks that:

- it opens without repair prompt;
- template presentation is materially intact;
- figures are visible and plausibly positioned;
- generated/prepared tables are intact;
- captions/numbering are current;
- semantic references display current values;
- included TOC/list structures display current entries;
- no unresolved YADG control text is visible;
- no obvious pagination/layout corruption was introduced.

This is a milestone completion gate, not perpetual re-review.

## Fixtures and hygiene

Use synthetic, redistribution-safe templates/content only.

Never commit confidential bank templates/content.

## Existing LibreOffice validation

M0005 LibreOffice behavior remains supported and historically validated.

M0009 does not require re-running its visual review unless implementation changes that renderer.

## Documentation changes

Implementation updates direct public usage docs when M0009 behavior would otherwise be missing or contradictory.

Broad V1 documentation audit remains M0010 work.

GitHub workflow automation is not M0009 scope.
