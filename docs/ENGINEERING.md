# Engineering

## Scope

This document defines repository-wide engineering and validation policy for YADG.

## Stack

YADG is implemented in C#/.NET.

The CLI is expected to use `System.CommandLine`.

DOCX authoring is based on OOXML, preferably through the Open XML SDK (`DocumentFormat.OpenXml`) unless implementation evidence requires a different package-level approach.

The Microsoft Word renderer/finalizer is a separate Windows-specific executable or component and may use Office Interop.

YADG is a product/tool repository, not a `dotnet-library` product. Internal class libraries are implementation structure, not a public package contract.

## Platform policy

### Office-independent path

Parsing, checking, semantic-model creation, and DOCX authoring must execute without Microsoft Office installed.

The implementation should remain portable across ordinary .NET environments to the extent supported by selected dependencies. Do not introduce Windows-only dependencies into the authoring components without changing architecture authority.

### Word-rendering path

Validation that depends on Microsoft Word targets:

- Windows;
- a locally installed compatible Microsoft Word/Office runtime;
- Office Interop capability;
- a non-interactive or controlled automation context suitable for the renderer's documented invocation.

The renderer must fail clearly when its required runtime is unavailable. A non-Word substitute is not evidence of Word behavior.

## Canonical engineering interface

The repository will expose canonical `eng/` commands or launchers as implementation proceeds. Implementation agents must use documented commands rather than inventing validation invocations.

M0001 is responsible for establishing the initial repository-specific engineering interface.

Cross-platform launchers may wrap the same repository logic, but complex product semantics must not live in shell/PowerShell wrappers.

## Test strategy

YADG uses **integration-first testing** where correctness materially depends on a boundary such as:

- Markdown parser behavior;
- OOXML package structure;
- Word run/container representation;
- DOCX relationships and fields;
- filesystem paths/assets;
- process boundaries;
- Microsoft Word rendering.

Unit tests are appropriate where isolated validation is materially cheaper, more exhaustive, or more diagnostic. Do not create unit tests merely to mirror implementation structure.

## Validation model

Validation depth and validation locus are independent.

### Tier 0 — edit sanity

Formatting, compilation of touched projects when cheap, schema/JSON sanity, and basic static checks.

### Tier 1 — focused validation

Narrow tests for the changed semantic or OOXML behavior.

### Tier 2 — repository validation

The normal Office-independent build and test suite for the repository.

Tier 2 must be runnable without Microsoft Word.

### Tier 3 — integration validation

Representative boundary validation against the concrete required target.

For ordinary authoring this includes real DOCX packages and prepared fixture templates, not mocks of OOXML behavior where package representation is material.

For Word-dependent renderer behavior the target is a real installed Microsoft Word runtime on Windows.

### Tier 4 — consumer/release validation

Exercise the distributable YADG artifact through its intended invocation mechanism rather than only invoking a project output directly.

Exact packaging/distribution is not yet fixed. A later release milestone must define the concrete consumer path.

### Tier 5 — human review

Use milestone-scoped human review when automation cannot establish whether generated Word/PDF output is visually and semantically acceptable.

Human review is evidence for the milestone that requires it, not a permanent requirement to re-review historical milestones after unrelated changes.

## DOCX validation

Do not use byte-for-byte DOCX equality as the primary correctness contract.

DOCX is a ZIP package and may contain volatile metadata, relationship identifiers, ordering differences, or normalization that does not change semantics.

Prefer structural assertions such as:

- expected Word element/style exists;
- expected heading/table/figure count or placement exists;
- target tag is resolved;
- unrelated template content remains;
- stable bookmark/reference/field structures resolve;
- package relationships point to valid parts;
- no forbidden unresolved YADG tags remain.

Golden documents may be used selectively for end-to-end regression evidence, especially after real rendering, but structural diagnostics remain the primary automated contract.

## Fixtures

Test fixtures should be intentionally small and representative.

Where Word-specific edge cases matter, include fixture documents that demonstrate them, including tags split across multiple OOXML runs.

Fixture provenance and purpose must be documented close to the fixture or test.

Do not commit confidential bank documents as fixtures.

## Determinism and provenance

Given equivalent source inputs, template inputs, configuration, and relevant tool versions, YADG should produce semantically equivalent authored artifacts.

Where volatile DOCX metadata prevents binary determinism, validation must explicitly ignore or normalize only the known volatile portions rather than weakening semantic assertions.

Generated evidence should identify enough provenance to determine what command/target produced it when retained for milestone review.

## Diagnostics

Validation failures should prefer actionable diagnostics over raw library exceptions.

User-facing failures should eventually carry stable codes where the class of failure is part of the supported product contract.

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

Implementation directly updates authority documents only when implementation would otherwise contradict them.

Broad documentation normalization is deferred to a separate documentation-sync pass.

`.guide-profile.json` and `.guide-sync/` are coordination metadata, not ordinary implementation authority.
