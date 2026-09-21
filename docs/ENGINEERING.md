# Engineering

## Scope

Repository-wide engineering and validation policy for YADG.

## Product boundary

YADG is C#/.NET.

Open XML authoring remains separate from real document renderers.

M0009 introduced a Windows/desktop-Microsoft-Word specialization and explicit publishing.

M0010 strengthens the pre-V1 compatibility boundary with realistic Office-produced templates, relative heading composition, and document-wide visible-text value substitution.

M0011's corrected release authority selects built-in late-bound Word COM for the V1 tool. This supersedes the earlier generated-interop packaging requirement for the release candidate.

## Platform policy after M0009

The authoritative full-product build and validation locus is Windows.

A complete Linux build is no longer a release/readiness requirement.

Authoring components should remain free of Word COM dependencies, but the CLI/full solution may be Windows-targeted because of the selected Word interop mechanism.

Do not add complexity solely to preserve Linux compilation if that conflicts with a robust Word renderer.

## Microsoft Word development/build capability

The authoritative Word-capable validation machine has:

- Windows;
- repository-supported .NET SDK;
- desktop Microsoft Word installed and COM-registered;
- an interactive user profile with Office activation/first-run complete.

The corrected M0011 Word renderer uses version-independent `Word.Application` late binding and does not require Office type-library generation tooling to build or package YADG.

## Word automation execution policy

Authoritative real Word execution occurs in an interactive logged-on Windows user session.

Do not claim support for automation from a Windows service, SYSTEM Task Scheduler job, ASP/ASP.NET host, DCOM server, or other non-interactive/server-side context.

Word tests/runs must avoid concurrent YADG-owned Word automation instances.

Tests must clean up owned documents/application instances and verify, where practical, that no owned Word process remains.

## LibreOffice capability

M0010 authoritative compatibility validation also requires a real LibreOffice Writer runtime capable of opening/finalizing the committed realistic fixtures through the existing LibreOffice renderer.

The existing isolated-profile/UNO runtime policy remains in force.

## Canonical engineering interface

The canonical repository validation command remains:

```powershell
./eng/validate.ps1
```

Its authoritative locus is Windows and it must build the Word-enabled product through the ordinary .NET SDK path; real Word remains required for Word-renderer execution.

Tier-3 realistic Office compatibility is separate from ordinary Tier-2 repository validation.

## Test strategy

YADG remains integration-first where real document/runtime behavior materially determines correctness.

Synthetic OpenXML fixtures remain appropriate for focused edge cases such as overflow, malformed story tags, precise run-splitting, and diagnostics.

Synthetic OpenXML-only fixtures are no longer sufficient evidence for broad template compatibility.

M0010 requires committed application-produced DOCX fixtures with recorded provenance.

## Application-produced fixture provenance

For each realistic fixture, commit repository-local provenance containing:

- origin application (`Microsoft Word` or `LibreOffice Writer`);
- exact application version;
- operating system/platform;
- creation/save method;
- SHA-256 of the committed DOCX;
- creation/update date;
- statement that the fixture is synthetic and redistribution-safe.

The fixture must originate as a new document in the stated application and its required content/structure must be created through that application's UI/document model before DOCX save.

Do not satisfy this requirement by synthesizing a DOCX with OpenXML SDK and merely round-tripping it through the application.

Tests verify the committed fixture hash against provenance and operate on copied fixture bytes rather than regenerating the fixture.

## Validation tiers

### Tier 0 — edit sanity

Use the repository's Windows/MSBuild-compatible build path and focused static/package checks.

### Tier 1 — focused

Cover deterministic authoring semantics without needing the external Office runtimes for every case:

- template outline-context resolution;
- `section` and `content` heading rebasing;
- selected roots whose source level is not 1;
- preservation of heading-level gaps;
- effective heading levels 1..9;
- failure before output mutation when effective level exceeds 9;
- front-matter heading bindings 1..9;
- effective-level style/numbering/reference validation;
- multi-paragraph body composition;
- value story discovery and split-run replacement;
- nested text-box scope isolation;
- footnote/endnote/comment/header/footer value diagnostics;
- field-instruction and metadata exclusions.

### Tier 2 — repository

Run:

```powershell
./eng/validate.ps1
```

Tier 2 must include focused M0010 tests and fixture-provenance/hash verification, but it need not launch Word/LibreOffice for the full compatibility matrix.

### Tier 3 — realistic Office compatibility

M0010 adds a dedicated Windows PowerShell integration target:

```powershell
./eng/test-m0010-tier3.ps1
```

The target requires:

```text
Windows interactive user session
Microsoft Word installed/activated/COM-registered
LibreOffice Writer installed
M0009 Word build prerequisites
```

It uses the committed Word-origin and LibreOffice-origin fixture templates and does not synthesize substitute templates.

It must run:

```text
Word-origin        -> YADG build -> Word render
Word-origin        -> YADG build -> LibreOffice render
LibreOffice-origin -> YADG build -> Word render
LibreOffice-origin -> YADG build -> LibreOffice render
```

For each path record:

- fixture identity/hash/origin;
- repository revision;
- OS/platform;
- renderer/runtime version;
- authored DOCX identity/hash;
- finalized DOCX identity/hash;
- command/result.

Automated inspection must verify at least:

- no unresolved expected value tags in supported stories;
- expected values present in body/header/footer/footnote/endnote/comment/text-box stories;
- correct rebased heading levels/styles for `section` and `content` scenarios;
- ordinary paragraphs preserve their template-owned paragraph style/properties;
- semantic references/numbering still target effective rendered headings correctly;
- existing fields/index structures remain present and renderer-updated where applicable;
- template-owned headers/footers/notes/comments/text boxes remain structurally present;
- existing static template content and representative layout structures remain present;
- finalized DOCX opens successfully in the real renderer path without repair/fatal conversion error.

No mock, generated-simple-template, or single-renderer substitute is authoritative Tier-3 evidence.

### Tier 4 — consumer/release

Deferred to M0011 V1 release-readiness.

### Tier 5 — human review

M0010 owns blocking artifact-quality review:

```text
HR-M0010-01
```

## M0010 human review

Review class:

```text
artifact-quality
```

Review subject consists of two representative finalized artifacts from the Tier-3 matrix:

1. Word-origin fixture finalized by Microsoft Word;
2. LibreOffice-origin fixture finalized by LibreOffice.

The review evidence also identifies the other two successful cross-render matrix results.

The reviewer checks the representative artifacts in their corresponding application and confirms:

- document opens without repair/conversion warning requiring intervention;
- page/layout/section structure remains materially intact;
- template-owned static content remains correctly positioned;
- headers and footers remain intact, including value substitution and page fields;
- footnote/endnote/comment values are visibly substituted;
- text-box/shape value is visibly substituted without destroying shape/layout;
- `section` and `content` insertion below template outline level 4 produces the expected visible heading hierarchy;
- inserted ordinary paragraphs use the intended body presentation and remain separate paragraphs;
- tables/figures/captions/references/TOC/list structures used by the fixture remain plausible/current;
- no unresolved YADG controls expected to be consumed are visible;
- no obvious pagination/layout corruption attributable to YADG is present.

The implementation extends existing milestone review tooling so:

```powershell
./eng/review-check.ps1 --milestone M0010
```

fails until the human approval record exists.

No implementation agent may fabricate approval.

Waiver is forbidden before V1 release-readiness.

## Fixtures and hygiene

All realistic and focused fixtures are synthetic and redistribution-safe.

Never commit confidential/bank templates or document content.

Tracked changes must be accepted in the realistic fixture templates before they become fixture authority.

## Existing M0009 validation

M0009 Word renderer/publish behavior remains in force.

M0010 may strengthen authoring logic consumed by both renderers, but must not weaken the M0009 Word integration/build boundary or publishing behavior.

## Documentation changes

Implementation updates direct README/user-facing documentation when M0010 behavior would otherwise be missing or contradictory.

The broad V1 documentation audit, NuGet packaging/publishing scripts, installation validation, version/release audit, and related release-readiness work move to M0011.

GitHub workflow automation remains postponed and is not M0010 scope.
