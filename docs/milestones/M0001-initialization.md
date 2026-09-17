# Milestone — M0001 Initial Implementation Substrate

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
| Focused validation | repository-specific M0001 focused command established during this milestone |
| Repository validation | canonical Office-independent repository validation command established during this milestone |
| Integration validation | real synthetic DOCX fixtures through OOXML authoring |
| Validation locus/platform | local or CI .NET environment; Microsoft Word not required |
| Consumer/release validation | not applicable; distributable packaging is out of scope |
| Human review | none required |

## Goal

Create the first executable YADG implementation substrate without weakening the settled architecture.

The milestone establishes a working .NET repository, a deterministic CLI/check path, a semantic Markdown-to-YADG model boundary, and a narrow but real DOCX template-tag transformation proving that visible textual tags can be located across OOXML run boundaries and replaced while preserving unrelated template structure.

This milestone deliberately stops before Word/Office rendering.

## Target State

When M0001 is complete:

- the repository builds from clean checkout using documented canonical `eng/` commands;
- YADG has a .NET solution and component structure consistent with `docs/ARCHITECTURE.md`;
- `yadg check` exists as an executable CLI path;
- Markdown can be parsed into a minimal semantic YADG model with explicit stable section IDs;
- a synthetic DOCX fixture containing a visible textual tag can be inspected through OOXML;
- tag recognition succeeds even when one logical tag is split across multiple Word runs;
- the implementation can resolve a section-content reference and replace the tag location in a copied DOCX fixture without Microsoft Word installed;
- focused structural assertions establish that replacement occurred and unrelated template structure/style was not globally flattened;
- malformed/unresolved references return actionable diagnostics and a failing process exit code;
- Office Interop is absent from the authoring dependency graph;
- Microsoft Word is not required for repository validation;
- implementation work and validation evidence are reconciled in `.execution/M0001-initialization.md` before closure.

## Scope

Implement enough product behavior to establish the architecture with real boundary behavior rather than placeholder interfaces.

Expected focus areas:

- .NET solution/repository setup;
- canonical `eng/` build/test/check launchers or commands;
- CLI composition using `System.CommandLine`;
- minimal semantic model for section references;
- Markdown parsing sufficient for the milestone fixture;
- stable explicit section IDs;
- minimal tag grammar sufficient to express section-content replacement;
- OOXML logical-text tag discovery across run boundaries;
- DOCX transformation on synthetic test fixtures;
- diagnostics and exit behavior for the implemented failure cases;
- structural tests against real DOCX packages.

The exact internal project names and local class design are implementation-owned as long as durable component boundaries remain recoverable and the architecture contract is preserved.

## Non-goals

M0001 does **not** implement:

- Microsoft Word/Office Interop rendering;
- PDF export;
- `yadg render`;
- `yadg publish`;
- full Markdown coverage;
- figures/images;
- production table population;
- captions, bookmarks, cross-reference rebuilding, TOCs, or list-of-figures behavior;
- workspace link/include semantics;
- alternative renderers;
- plugin loading;
- DMS integration;
- public NuGet/library packaging;
- final release packaging;
- migration of confidential real-world bank templates into test fixtures.

Do not add speculative abstractions for these features unless required to keep an already-set architectural boundary intact.

## Decisions and Constraints

- YADG is a product/tool repository, not a `dotnet-library` product.
- Markdown content and prepared Word templates have the ownership defined in `docs/SPECS.md`.
- The authoring path must not require Microsoft Office.
- Microsoft Word/Office Interop dependencies are forbidden in M0001 authoring projects.
- Template tags are visible text, not Word content controls.
- Tag discovery must tolerate a logical tag split across multiple OOXML runs.
- Stable semantic IDs are required for Markdown references.
- For heading-based references, complete section selection and section-content selection are distinct semantics. Do not implement this as generic `append`/`replace`.
- The minimal implemented M0001 grammar may support only the section-content form required by the fixture, provided its naming/shape does not contradict the full semantic model.
- The semantic Markdown model must not contain Open XML SDK element types.
- Tests must exercise real DOCX package behavior for OOXML-sensitive correctness.
- Do not use binary DOCX equality as the main acceptance test.
- Do not require unit tests by default; choose them only where they improve diagnostic value or coverage.
- Synthetic, redistribution-safe fixtures are mandatory because the repository is public.

## Baseline Executor Readiness

All architecture, ownership, profile, test-policy, and M0001 acceptance decisions required for execution are contained in this milestone and the authority documents below.

Implementation freedom includes:

- solution/project names;
- local namespaces and class decomposition;
- choice of Markdown parser package if compatible with the contract;
- exact minimal textual tag spelling for M0001, provided it clearly represents section-content selection and remains easy to extend;
- exact `eng/` launcher implementation;
- test framework;
- fixture construction method.

A choice becomes a planning issue only if it changes project semantics, component ownership, public compatibility, or validation policy.

## Execution Tractability

At implementation start, create:

```text
.execution/M0001-initialization.md
```

Use it to record bounded work packages, coverage of milestone obligations, validation commands/evidence, blockers, and closure reconciliation.

Do not pre-plan line-by-line edits. Keep work packages coherent enough that each can be validated before moving on.

## Required Authority

Read these files before implementation:

- `docs/SPECS.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`

Do not read the external guide repository to recover implementation rules.

## Acceptance Criteria

- A clean checkout can execute the documented canonical repository build and test/check flow with the supported .NET SDK after initial restore.
- The solution clearly separates Office-independent semantic authoring from Word-specific OOXML authoring.
- No Office Interop package/reference is present in the authoring dependency graph.
- An executable `yadg check` command exists and returns success for the valid milestone fixture.
- The Markdown fixture contains an explicit stable section ID and resolves to a semantic section node.
- Section-content selection excludes the referenced heading and includes content that structurally belongs to it.
- A synthetic DOCX fixture contains a visible textual tag whose characters are split across at least two OOXML runs.
- YADG recognizes the split tag as one logical tag.
- YADG produces a copied/transformed DOCX whose target tag location contains the expected resolved Markdown content.
- Structural assertions demonstrate that unrelated paragraphs and styles in the template remain present and are not globally reformatted.
- At least one invalid fixture demonstrates an unresolved or malformed reference with an actionable diagnostic and non-zero CLI exit code.
- Repository validation requires no installed Microsoft Word.
- No confidential or proprietary fixture content is committed.
- `.execution/M0001-initialization.md` is complete enough to map every acceptance criterion to implementation and validation evidence before the milestone is marked complete.

## Validation

### Tier 0 — edit sanity

Use the canonical engineering interface established in this milestone to verify solution/configuration syntax and compilation of affected projects.

Expected evidence: command exit status and compiler/static-check output.

### Tier 1 — focused semantic and OOXML validation

Run focused tests covering:

- section/content selection boundaries;
- stable reference resolution;
- logical tag reconstruction across multiple OOXML runs;
- replacement of the split-run tag;
- unresolved/malformed reference diagnostics.

Target: the actual chosen Markdown parser and actual OOXML/Open XML SDK representation.

Expected evidence: passing focused tests and structurally inspected synthetic DOCX output.

### Tier 2 — repository validation

Run the canonical Office-independent repository validation command established by M0001.

It must restore/build/test the repository sufficiently to establish ordinary authoring correctness without invoking Microsoft Word.

Expected evidence: zero exit status and complete test result.

### Tier 3 — integration validation

Target: real DOCX ZIP/OOXML package fixtures processed through the implemented YADG authoring path.

Execution locus: any supported local/CI environment with the .NET SDK; Microsoft Word must not be required.

Invocation: use the repository-owned integration test/fixture command established during implementation.

Evidence must include structural assertions against produced DOCX content. Do not substitute a mocked OOXML object graph if the tested behavior depends on actual package/run structure.

### Tier 4 — consumer/release validation

Not applicable to M0001. Packaging/distribution is intentionally unresolved.

### Tier 5 — human review

Not required. M0001 acceptance is structurally automatable; visual Word rendering is outside scope.

## Documentation Policy

Update project authority only if implementation discovers a direct contradiction that must be resolved for M0001.

Do not broaden M0001 into a documentation cleanup pass.

If documentation changes are desirable but not required for correctness, write a deferred synchronization hint under `.guide-sync/pending/` rather than blocking milestone execution.

## Completion Expectations

Implementation owns milestone closure:

```text
execution decomposition
-> persistent execution ledger
-> implement/validate bounded work packages
-> freshly reread M0001 and required authority
-> reconcile milestone <-> ledger <-> repository/evidence
-> completion audit
```

Passing tests alone does not establish completion if any acceptance criterion is unmapped or contradicted.

## Escalation Boundary

Return to planning if implementation requires a new decision that materially changes:

- the Markdown/template ownership model;
- textual-tag semantics;
- stable reference semantics;
- authoring/rendering component boundaries;
- the Office-independent requirement;
- M0001 scope or acceptance criteria;
- validation targets or policy.

Local code structure, package choices consistent with the contract, test mechanics, and execution sequencing remain implementation-owned.
