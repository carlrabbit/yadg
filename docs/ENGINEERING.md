# Engineering

## Scope

This document defines repository-wide engineering and validation policy for YADG.

## Stack

YADG is implemented in C#/.NET.

The CLI uses `System.CommandLine`.

DOCX authoring is based on OOXML through the Open XML SDK (`DocumentFormat.OpenXml`) unless later implementation evidence produces a project-level decision to change that boundary.

The Microsoft Word renderer/finalizer is a separate Windows-specific executable or component and may use Office Interop.

YADG is a product/tool repository, not a `dotnet-library` product. Internal class libraries are implementation structure, not a public package contract.

## Platform policy

### Office-independent path

Parsing, checking, semantic-model creation, workspace discovery, and DOCX authoring must execute without Microsoft Office installed.

The implementation should remain portable across ordinary .NET environments to the extent supported by selected dependencies. Do not introduce Windows-only dependencies into the authoring components without changing architecture authority.

### Word-rendering path

Validation that depends on Microsoft Word targets:

- Windows;
- a locally installed compatible Microsoft Word/Office runtime;
- Office Interop capability;
- a non-interactive or controlled automation context suitable for the renderer's documented invocation.

The renderer must fail clearly when its required runtime is unavailable. A non-Word substitute is not evidence of Word behavior.

M0002 does not invoke this path.

## Canonical engineering interface

The canonical repository validation command established by M0001 is:

```powershell
./eng/validate.ps1
```

Repository-specific focused commands may live under `eng/`, but product semantics must remain in product code rather than PowerShell/shell wrappers.

Implementation agents must use documented repository commands rather than inventing alternative validation interfaces when an applicable canonical command exists.

## Test strategy

YADG uses **integration-first testing** where correctness materially depends on a boundary such as:

- Markdown parser behavior;
- workspace/filesystem discovery;
- OOXML package structure;
- Word run/container representation;
- DOCX relationships and fields;
- process boundaries;
- Microsoft Word rendering.

Unit tests are appropriate where isolated validation is materially cheaper, more exhaustive, or more diagnostic. Do not create unit tests merely to mirror implementation structure.

## Validation model

Validation depth and validation locus are independent.

### Tier 0 — edit sanity

Formatting, compilation of touched projects when cheap, schema/JSON sanity, and basic static checks.

### Tier 1 — focused validation

Narrow tests for changed semantic, workspace, Markdown, or OOXML behavior.

### Tier 2 — repository validation

The normal Office-independent repository validation command is:

```powershell
./eng/validate.ps1
```

Tier 2 must remain runnable without Microsoft Word.

### Tier 3 — integration validation

Representative boundary validation against the concrete required target.

For ordinary authoring this includes real filesystem workspaces, real Markdown parsing, and real DOCX packages/prepared fixture templates rather than mocks where those representations materially determine correctness.

For Word-dependent renderer behavior the target is a real installed Microsoft Word runtime on Windows.

### Tier 4 — consumer/release validation

Exercise the distributable YADG artifact through its intended invocation mechanism rather than only invoking a project output directly.

Exact packaging/distribution is not yet fixed. A later release milestone must define the concrete consumer path.

### Tier 5 — human review

Use milestone-scoped human review when automation cannot establish whether generated Word/PDF output is visually and semantically acceptable.

Human review is evidence for the milestone that requires it, not a permanent requirement to re-review historical milestones after unrelated changes.

M0002 does not require human review because its acceptance contract is structural and semantic, not layout-rendering acceptance.

## DOCX validation

Do not use byte-for-byte DOCX equality as the primary correctness contract.

DOCX is a ZIP package and may contain volatile metadata, relationship identifiers, ordering differences, or normalization that does not change semantics.

Prefer structural assertions such as:

- expected Word element/style exists;
- expected heading/paragraph placement exists;
- target tag is resolved;
- unrelated template content remains;
- package relationships point to valid parts;
- no forbidden unresolved YADG tags remain.

Golden documents may be used selectively for end-to-end regression evidence, especially after real rendering, but structural diagnostics remain the primary automated contract.

## Workspace/filesystem validation

Workspace behavior must be tested against real temporary directory trees where discovery semantics are material.

M0002 validation must cover:

- root-marker discovery;
- recursive Markdown discovery/exclusions;
- nested-workspace rejection;
- template/output conventions;
- symlink/reparse-point rejection where the test platform can create the relevant filesystem object.

When the current platform cannot create a particular link/reparse-point form without additional privileges, the unavailable case does not block portable Tier 2 validation. Platform-capable focused validation should cover it, and implementation must not weaken the project rule merely because one CI locus cannot construct the fixture.

## Fixtures

Test fixtures should be intentionally small and representative.

Where Word-specific edge cases matter, include fixture documents that demonstrate them, including tags split across multiple OOXML runs.

Fixture provenance and purpose must be documented close to the fixture or test.

Do not commit confidential bank documents as fixtures.

## Determinism and provenance

Given equivalent source inputs, template inputs, configuration/conventions, and relevant tool versions, YADG should produce semantically equivalent authored artifacts.

Where volatile DOCX metadata prevents binary determinism, validation must explicitly ignore or normalize only the known volatile portions rather than weakening semantic assertions.

Generated evidence should identify enough provenance to determine what command/target produced it when retained for milestone review.

## Diagnostics

Validation failures should prefer actionable diagnostics over raw library exceptions.

User-facing failures should carry stable codes where the class of failure is part of the supported product contract.

Established diagnostic codes must not be silently reassigned to incompatible meanings.

## Dependency constraints

Dependencies must preserve the architectural boundary:

- core authoring must not reference Office Interop;
- renderer-specific dependencies must not leak into portable authoring;
- Word-specific OOXML representations must not become the semantic Markdown model.

## Public repository hygiene

The repository is public.

Do not commit:

- confidential bank templates or documentation;
- customer/internal identifiers;
- credentials;
- proprietary example content;
- files whose redistribution rights are unclear.

Synthetic fixtures must be used for public tests and examples.

## Documentation changes

Implementation directly updates authority or public documentation when implemented behavior would otherwise contradict it.

Broad documentation normalization is deferred to a separate documentation-sync pass.

`.guide-profile.json` and `.guide-sync/` are coordination metadata, not ordinary implementation authority.
