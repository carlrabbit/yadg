# Engineering

## Scope

This document defines repository-wide engineering and validation policy for YADG.

## Stack and product boundary

YADG is implemented in C#/.NET. The CLI uses `System.CommandLine`. Office-independent DOCX authoring uses OOXML through the Open XML SDK.

YADG is a product/tool repository, not a reusable `dotnet-library` product merely because it contains class-library projects.

## Office-independent authoring

Parsing, checking, semantic-model creation, workspace discovery, and `build` must execute without Microsoft Office or LibreOffice installed.

Renderer dependencies must not leak into `authoring-core` or `word-authoring`.

## LibreOffice renderer specialization

M0005 defines LibreOffice as the first real renderer target.

Authoritative renderer integration validation targets a locally installed LibreOffice runtime capable of:

- headless/API automation;
- opening the M0004-authored DOCX representation;
- refreshing text fields;
- updating Writer document indexes;
- saving DOCX through LibreOffice's Office Open XML Writer filter;
- exporting PDF through `writer_pdf_Export`.

The execution locus is local or CI only when that locus actually provides the required LibreOffice runtime. Portable tests are not substitutes for the real renderer target.

Each render invocation uses a unique temporary LibreOffice user profile and a local-only programmatic connection. It must not use the user's normal profile or attach to an unrelated running instance.

The concrete LibreOffice version and operating system are validation provenance. M0005 uses capability-based validation rather than claiming untested versions/platforms by number alone.

If LibreOffice is unavailable at the current locus, Office-independent Tier 0-2 validation may still run, but Tier 3 renderer success must not be claimed.

## Canonical engineering interface

The canonical repository validation command is:

```powershell
./eng/validate.ps1
```

Product semantics remain in product code rather than shell wrappers.

## Integration-first testing

Prefer the largest practical real boundary where correctness depends on Markdown parsing, filesystem discovery, OOXML/DOCX structure, image representation, process boundaries, or renderer behavior.

Use unit tests when isolated validation is substantially cheaper or more diagnostic; do not create a test pyramid by policy.

## Validation tiers

### Tier 0 — edit sanity

Compilation/static/schema checks for changed areas.

### Tier 1 — focused validation

Focused semantic, authoring, renderer-client/process, review-workflow, and failure-path checks.

### Tier 2 — repository validation

```powershell
./eng/validate.ps1
```

Tier 2 must remain runnable without LibreOffice or Microsoft Word unless a later repository-wide policy explicitly changes that contract.

### Tier 3 — integration validation

For authoring: real filesystem workspaces and real DOCX packages.

For M0005 rendering: a real installed LibreOffice process, real M0004-authored DOCX inputs, real field/index refresh, real DOCX save, and real PDF export.

Mocks, a fake process, or direct OOXML manipulation are not equivalent evidence for renderer behavior.

### Tier 4 — consumer/release validation

Deferred until packaging/distribution is fixed.

### Tier 5 — human review

Use milestone-scoped human review when automation cannot judge visual/artifact quality.

M0005 owns blocking review `HR-M0005-01`.

## M0005 human-review workflow

M0005 introduces the repository's human-review substrate conforming to the external guide's milestone-scoped review model. Project-local durable state is:

```text
.review/pending/
.review/closed/
.review/records/
```

Generated review evidence is local/derived state under:

```text
artifacts/review/evidence/
artifacts/review/session/
```

The canonical M0005 review-check command is:

```powershell
./eng/review-check.ps1 --milestone M0005
```

Implementation must also provide repository-local human commands sufficient to list/show the pending review and record an explicit human decision without permitting an implementation agent to fabricate approval. PowerShell is the canonical launcher; POSIX equivalents may mirror it consistently with repository launcher policy.

`review-check` validates only the explicit milestone context. It must fail while a blocking M0005 review is pending/changes-requested/rejected and pass only after an acceptable recorded decision.

Completed review records are historical milestone evidence and are not invalidated by future commits.

## DOCX validation

Do not use byte-for-byte DOCX equality as the primary contract. Prefer structural assertions for elements, styles, fields, bookmarks, relationships, preserved template content, and absence of unresolved YADG controls.

After M0005, renderer validation additionally checks observable evaluated values in the finalized DOCX/PDF against the real renderer target.

## Workspace/filesystem validation

Use real temporary directory trees where discovery semantics matter. Preserve existing path/symlink/reparse-point rules.

## Fixtures and public hygiene

Fixtures must be small, synthetic, redistribution-safe, and document their purpose where Word/LibreOffice-specific edge cases matter.

Never commit confidential bank templates/content, credentials, customer/internal identifiers, or assets with unclear redistribution rights.

## Determinism and provenance

Equivalent source/template inputs should produce semantically equivalent authoring output. Renderer output may contain engine/version-dependent volatile state; validation must distinguish that from semantic/rendered correctness.

Retained or reviewed renderer evidence identifies the LibreOffice version/platform and enough input provenance to reproduce the result.

## Diagnostics

Prefer actionable product diagnostics over raw library/process exceptions. Established diagnostic codes must not be silently repurposed.

## Documentation changes

Implementation updates project authority/public docs when implemented behavior would otherwise contradict them. Broad documentation synchronization remains a separate pass.

`.guide-profile.json` and `.guide-sync/` are coordination metadata, not ordinary implementation authority.
