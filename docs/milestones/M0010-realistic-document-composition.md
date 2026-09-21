# Milestone — M0010 Realistic Document Composition and Template Compatibility

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
| Implementation autonomy | high inside the fixed contracts below |
| Documentation sync | deferred; direct contradictions fixed in this overlay |
| Repository validation | `./eng/validate.ps1` on authoritative Windows build locus |
| Integration validation | real application-produced fixtures + real Word + real LibreOffice |
| Tier-3 command | `./eng/test-m0010-tier3.ps1` |
| Validation locus | interactive Windows user session with Word and LibreOffice installed |
| Consumer/release validation | deferred to M0011 |
| Human review | blocking `HR-M0010-01` |

## Goal

Before starting V1 release-readiness, replace YADG's overly simple compatibility proofs with realistic document-composition evidence and close two product gaps exposed by those proofs:

1. compose Markdown headings/paragraphs correctly inside an existing template heading hierarchy;
2. make workspace value replacement behave like document-wide visible template-text substitution across the Word stories a normal template author can edit.

The milestone must prove these behaviors on committed DOCX templates originating from both Microsoft Word and LibreOffice Writer, not only on OpenXML-synthesized fixtures.

## Target State

A template author can prepare a realistic DOCX such as:

```text
[cover/layout/static content]
Heading level 4 — Template-owned section
    {{section:architecture}}

[headers / footers / notes / comments / text boxes]
Document version {{value:document-version}}
```

with Markdown such as:

```markdown
# Architecture {#architecture}

First real paragraph.

Second paragraph with **strong** and *emphasized* content.

## Assumptions

Details.
```

and YADG authors:

```text
Heading level 4 — Template-owned section
    effective Heading level 5 — Architecture
        body paragraph
        body paragraph
        effective Heading level 6 — Assumptions
```

while substituting the workspace value in every supported visible Word story without the template author needing to know which OOXML part stores that text.

The resulting authored documents from both a Word-origin and LibreOffice-origin template can be finalized successfully by both supported renderers.

## Scope

### Realistic application-produced template fixtures

- one committed Microsoft-Word-origin DOCX template;
- one committed LibreOffice-Writer-origin DOCX template;
- fixture provenance including application/version/platform/hash;
- realistic layout, sections, static content, headings, fields, headers/footers, notes/comments/text boxes, tables/figures/layout structures;
- tests consume immutable committed fixture bytes rather than recreating the templates.

### Relative heading composition

- derive template heading context from effective Word outline level, not style names;
- rebase selected Markdown section hierarchy for both `section` and `content` placement;
- preserve relative source-level gaps;
- support effective Word heading levels 1 through 9;
- extend template front-matter heading style bindings to 1 through 9;
- use effective output level/style for heading-style validation, numbering, bookmarks/numeric section-reference eligibility, and renderer-visible outline structures;
- validate overflow before output mutation.

### Real paragraph composition

- multiple Markdown paragraphs remain distinct Word paragraphs;
- ordinary paragraphs preserve the placement anchor's template-owned paragraph properties/style;
- existing inline semantics continue to work in realistic templates;
- surrounding static template content is preserved.

### Document-wide visible value substitution

Support `{{value:<id>}}` in ordinary visible text within:

- main body/table cells;
- headers and footers, including variants;
- footnotes;
- endnotes;
- comments;
- Wordprocessing text-box/shape text nested in supported stories;
- ordinary visible text wrapped by hyperlinks/content controls.

Preserve split-run replacement and first-tag-character formatting semantics.

### Compatibility proof

Run both realistic templates through YADG build and both renderers, producing a four-path matrix.

Perform automated structural/semantic evidence and a blocking human visual/artifact review.

## Non-goals

M0010 does not add Markdown syntax or semantic creation APIs for:

- footnotes;
- endnotes;
- comments;
- text boxes/shapes;
- headers;
- footers.

M0010 does not move `section`, `content`, `table`, `figure`, or prepared-table block placement into those stories.

M0010 does not add:

- generic arbitrary-OOXML text substitution;
- tracked-change authoring semantics;
- new Markdown block types;
- new image formats;
- new renderer IDs;
- renderer equality guarantees;
- PDF publication;
- NuGet/.NET tool packaging;
- NuGet publishing scripts;
- V1 version/release audit;
- GitHub workflows/releases.

Those release-readiness concerns move to M0011.

## Resolved Decisions and Constraints

### Heading context source

Heading context is a prepared-template property, not a Markdown/parser property.

For every `section`/`content` placement, determine the nearest preceding top-level main-body template paragraph with an effective Word outline level 1 through 9.

Resolve outline level from actual paragraph/style semantics, including style inheritance.

Do not infer context from style names.

Do not let headings generated by earlier YADG placements influence context for a later placement.

If no such template heading exists, context level is 0.

### Heading mapping formula

Let:

```text
C = template context level
R = selected Markdown section root source level
L = emitted heading source level
```

For `{{section:id}}`:

```text
effective level = C + 1 + (L - R)
```

For headings emitted from `{{content:id}}`:

```text
effective level = C + (L - R)
```

The `content` root itself remains omitted.

These formulas are project policy. Implementation must not choose a different normalization/offset strategy.

### Heading depth and style roles

Effective levels are limited to 1 through 9.

A required level greater than 9 is a validation error before normal output changes.

Template front-matter version remains 1 and accepts heading bindings `1` through `9`.

Default heading roles are `Heading1` through `Heading9`.

Only required effective styles need to exist.

### Paragraph inheritance

Ordinary generated Markdown paragraphs continue to use the placement anchor as the paragraph-format/style prototype.

Heading paragraphs use their mapped effective heading role instead.

Do not introduce a new Markdown body-style configuration merely for M0010.

### Value story semantics

Workspace values are visible-template-text substitution, not block authoring.

Supported story/part set is fixed by `docs/specs/DOCUMENT-COMPOSITION.md`.

A value tag is resolved within one ordinary paragraph/text-container scope and may span runs, but it cannot span separate paragraphs, parts, or nested text boxes.

Nested text-box paragraphs are independent scopes.

The same supported-story enumeration must be used by `check` validation and `build` replacement.

### Value exclusions

Do not modify field instruction code, properties, relationship targets, bookmark names, custom XML, package metadata, alternative-text attributes, or arbitrary XML attributes.

Do not invent special Markdown/story creation syntax.

### Fixture origin

The realistic templates are implementation artifacts, but their provenance rule is fixed:

- each begins as a new document in its stated originating application;
- required structure/content is created through that application UI/document model;
- that application saves the DOCX;
- OpenXML SDK synthesis followed by resave is not sufficient provenance.

Exact fixture filenames/paths and creation automation mechanics are implementation-owned.

### Compatibility matrix

All four matrix paths are mandatory Tier-3 evidence:

```text
Word origin -> Word renderer
Word origin -> LibreOffice renderer
LibreOffice origin -> Word renderer
LibreOffice origin -> LibreOffice renderer
```

No one path substitutes for another.

### M0009 preservation

M0010 must preserve:

- M0009 MSBuild/COMReference-generated Word interop build path;
- Word renderer behavior/locus;
- publish behavior;
- LibreOffice existing output behavior.

M0010 may refactor authoring code as necessary but must not reopen the Word interop approach or renderer/publish contracts.

## Baseline Executor Readiness

Planning has fixed the material decisions that implementation must not invent:

- application-produced fixture requirement and provenance;
- exact realistic compatibility matrix;
- heading-context definition;
- exact `section`/`content` level formulas;
- effective-level range and overflow semantics;
- heading style-binding extension to levels 1..9;
- paragraph inheritance rule;
- supported value-story set;
- value replacement scope and exclusions;
- no special Markdown story syntax;
- validation loci/tiers;
- human-review subject/gate;
- M0011 release-readiness boundary.

Implementation owns concrete types/functions, OpenXML traversal mechanics, style-chain resolution mechanics, fixture file layout, fixture creation automation/manual mechanics consistent with provenance, focused test decomposition, diagnostics codes/messages, and implementation sequence.

## Execution Tractability

Create and maintain:

```text
.execution/M0010-realistic-document-composition.md
```

Decompose the milestone into bounded work packages and map every acceptance criterion to concrete implementation/validation evidence.

Do not keep coverage only in conversational memory.

## Required Authority

Read before implementation:

- `docs/SPECS.md`
- `docs/specs/DOCUMENT-COMPOSITION.md`
- `docs/specs/WORKSPACE-VALUES.md`
- `docs/specs/STRUCTURED-CONTENT.md`
- `docs/specs/WORD-REFERENCES.md`
- `docs/specs/RENDERING.md`
- `docs/specs/WORD-RENDERER.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`
- `.review/pending/HR-M0010-01.md`

Do not require the implementation agent to read the external guide repository, `.guide-profile.json`, planning conversation, or prior research to reconstruct the contract.

## Acceptance Criteria

### Realistic fixture authority

- repository contains one Word-origin and one LibreOffice-origin committed DOCX template fixture;
- each fixture has repository-local provenance with origin application, exact version, platform, creation/save method, date, SHA-256, and synthetic/non-confidential statement;
- committed fixture hash matches provenance;
- fixture content/structure originated in the stated application rather than OpenXML synthesis followed by resave;
- compatibility tests copy/use committed fixture bytes rather than recreating them;
- fixtures contain the required realistic content from `docs/specs/DOCUMENT-COMPOSITION.md`;
- tracked changes are accepted before fixture commit.

### Template outline context

- context is based on the nearest preceding top-level prepared-template body paragraph with effective outline level;
- paragraph-direct outline level works;
- style-provided/inherited outline level works;
- custom style names work without `HeadingN` name inference;
- intervening static non-heading paragraphs do not change the context;
- absence of a prior outline heading yields context 0;
- generated headings from another placement do not affect context analysis.

### `section` heading rebasing

- selected root is emitted one effective level below template context;
- descendants preserve their source-level delta from selected root;
- a selected source root whose source level is not 1 is normalized relative to itself, not treated as absolute Word level;
- skipped Markdown level gaps remain skipped after rebasing;
- effective style is selected from the mapped level, not the source level.

### `content` heading rebasing

- selected root remains omitted;
- a direct Markdown child of the omitted root is emitted one effective level below template context;
- deeper descendants preserve their relative source-level delta;
- ordinary root body paragraphs are emitted directly below the template context without a synthetic heading;
- skipped Markdown level gaps remain preserved.

### Heading level/style boundaries

- template front matter accepts heading style bindings `1..9` under existing version 1;
- default roles exist for `Heading1` through `Heading9`;
- only actually required effective heading styles are required;
- effective level greater than 9 fails `check`/`build` before normal output mutation;
- missing required effective style fails before output mutation;
- numeric section target/numbering validation uses the effective level/style;
- rebased headings remain usable by template-owned TOC/index update through both renderers.

### Real paragraph composition

- at least two consecutive Markdown paragraphs become distinct Word paragraphs in each realistic fixture;
- ordinary inserted paragraphs retain the intended anchor/body paragraph style/properties;
- heading styling is not leaked into ordinary paragraphs;
- emphasis/strong/break/reference behavior remains correct;
- static template paragraphs before/between/after placements survive unchanged in semantic content;
- existing table/figure/list/caption behavior remains compatible.

### Supported value stories

For both realistic fixtures, value substitution succeeds in expected ordinary visible text in:

- main body/table cell;
- header;
- footer;
- footnote;
- endnote;
- comment;
- text box/shape text.

At least one fixture also proves a first/even-page header/footer variant and a text box nested outside the ordinary main-body paragraph flow.

### Value replacement correctness

- split-run value tag works in a non-main-body story;
- replacement inherits formatting at the first logical tag character;
- surrounding static text/formatting survives;
- hyperlink/content-control wrappers do not by themselves disable ordinary visible-text value replacement;
- nested text-box text is processed separately from its containing outer paragraph;
- shared header/footer part is not duplicated merely because multiple sections reference it;
- literal/non-recursive semantics remain unchanged;
- missing/malformed supported-story values fail during `check` and `build` before output changes;
- no expected supported-story value tag remains unresolved after successful build.

### Value exclusions

- field instruction code is not altered;
- document/custom properties and package metadata are not altered;
- relationship targets/bookmark names/custom XML/alternative-text attributes are not treated as value targets;
- no new generic arbitrary XML substitution surface is introduced.

### Realistic four-path compatibility

All four required Tier-3 paths complete successfully using committed realistic fixtures:

- Word-origin -> Word renderer;
- Word-origin -> LibreOffice renderer;
- LibreOffice-origin -> Word renderer;
- LibreOffice-origin -> LibreOffice renderer.

For all four:

- authored/finalized DOCX is structurally readable;
- expected supported-story values are present;
- expected rebased heading structure is present;
- representative existing template structures remain present;
- renderer finalization succeeds without repair/fatal conversion failure;
- field/index/reference structures used by the fixture remain valid/current according to renderer-specific observable evidence.

### Regression boundaries

- existing M0002-M0009 validation remains green unless a test encoded the superseded absolute-heading/value-location limitation, in which case it is deliberately updated to the M0010 contract;
- existing LibreOffice behavior remains supported;
- existing Word renderer COMReference build/runtime boundary remains supported;
- `publish` behavior is unchanged;
- no new PDF publication contract is introduced.

### Documentation and hygiene

- direct README/user documentation explains relative heading composition with at least one `Heading4 + Markdown H1 -> effective Heading5` example;
- direct documentation states values work across supported visible Word stories without requiring users to understand OOXML parts;
- direct documentation clearly says Word structures such as footnotes/text boxes remain template-owned and are not created by new Markdown syntax;
- no confidential fixture content is committed;
- execution ledger maps each criterion to evidence.

### Human review

- implementation prepares the two representative finalized artifacts and complete Tier-3 evidence for `HR-M0010-01`;
- review tooling recognizes M0010 without weakening historical M0005/M0009 review checks;
- pending/negative/missing review causes `./eng/review-check.ps1 --milestone M0010` to fail;
- implementation does not fabricate approval;
- waiver is unavailable;
- milestone cannot become complete until actual human approval is recorded.

## Validation

### Tier 0 — edit sanity

Use the Windows/MSBuild-compatible repository build path already established by M0009.

Do not replace the generated Word interop build with another strategy.

### Tier 1 — focused authoring validation

Focused tests must cover the deterministic semantic/OOXML cases listed in `docs/ENGINEERING.md`, including synthetic edge cases for:

- heading context/style inheritance;
- both mapping formulas;
- non-H1 selected roots;
- level gaps;
- effective levels through 9 and overflow;
- value story enumeration;
- split-run/nested-text-box replacement scope;
- unsupported metadata/instruction surfaces.

These tests may use targeted synthetic OpenXML fixtures because they are not claiming broad application-template compatibility.

### Tier 2 — repository validation

Run:

```powershell
./eng/validate.ps1
```

Required evidence:

- full Word-enabled solution builds through the existing M0009-compatible Windows build path;
- focused M0010 tests pass;
- fixture provenance/hash verification passes;
- prior repository validation remains green subject only to deliberate contract updates.

Tier 2 does not claim real Word/LibreOffice matrix success.

### Tier 3 — realistic application-produced compatibility

Create/maintain the explicit integration command:

```powershell
./eng/test-m0010-tier3.ps1
```

Required locus:

```text
interactive Windows user session
Microsoft Word installed/activated/COM-registered
LibreOffice Writer installed
Visual Studio/MSBuild Word interop build prerequisites
```

The command must:

1. verify committed fixture hashes/provenance;
2. create temporary workspaces by copying the committed fixture bytes;
3. supply synthetic Markdown/YADG.md/assets required by the scenario;
4. run real `check` and `build` for both fixtures;
5. inspect authored DOCX for heading/story/value assertions;
6. run the four renderer matrix paths;
7. inspect finalized DOCX/renderer-observable results;
8. record the evidence required by `docs/ENGINEERING.md`;
9. preserve/copy the two representative review artifacts/evidence into the repository's existing review-evidence convention when requested by the implementation/review workflow;
10. return non-zero if any matrix path or required assertion fails.

Do not regenerate simpler fallback templates when a fixture path fails.

Do not omit one renderer/origin combination and still claim Tier-3 success.

### Tier 4 — consumer/release

Deferred to M0011.

### Tier 5 — blocking human review

Review ID:

```text
HR-M0010-01
```

Follow:

```text
.review/pending/HR-M0010-01.md
```

Completion requires actual human approval and:

```powershell
./eng/review-check.ps1 --milestone M0010
```

passing against that record.

## Constrained Runtime

The implementation/validation environment may be missing one of the two required desktop Office runtimes or application-creation capabilities.

If Microsoft Word or LibreOffice is unavailable:

- complete portable/focused implementation and Tier 0-2 evidence that is still valid;
- do not synthesize a substitute "realistic" fixture with OpenXML SDK;
- do not substitute one renderer for the missing runtime;
- do not claim Tier-3 completion;
- record the milestone as externally blocked on the missing declared capability.

If an application cannot create/save one of the required realistic fixture structures in DOCX while satisfying the contract, return to planning rather than silently deleting that story from the proof.

Human approval remains external to the implementation agent.

## Research

No durable research artifact is required.

The milestone is driven by observed repository/test limitations and explicit product decisions. All implementation-affecting conclusions are promoted directly into project authority in this overlay.

Research remains inactive under the current guide profile.

## Documentation Impact

Planning adds/updates:

- `docs/specs/DOCUMENT-COMPOSITION.md`;
- `docs/specs/WORKSPACE-VALUES.md`;
- `docs/SPECS.md`;
- `docs/ARCHITECTURE.md`;
- `docs/ENGINEERING.md`;
- `docs/TERMINOLOGY.md`;
- `docs/MILESTONES.md`;
- `.review/pending/HR-M0010-01.md`;
- this milestone.

Implementation updates direct README/user-facing usage documentation and review tooling as required by acceptance.

Broad documentation audit/release-readiness is not performed here; it moves to M0011 together with NuGet packaging/publishing scripts and consumer installation validation.

No `.guide-sync/pending/` hint is required because the release-readiness boundary is explicit and direct contradictions are resolved in this overlay.

No GitHub workflow work is included.

## Human Review

Applicability: `blocking`

Review class: `artifact-quality`

Review ID: `HR-M0010-01`

Owning milestone: `M0010`

Waiver: forbidden.

Review request:

```text
.review/pending/HR-M0010-01.md
```

The implementation agent prepares evidence/artifacts but only a human reviewer records the decision.

## Completion Expectations

Implementation owns:

```text
read milestone and Required Authority
-> create/reconcile execution ledger
-> derive bounded work packages
-> implement heading composition/value-story support
-> create and commit provenance-backed Word/LibreOffice fixtures
-> focused validation
-> ./eng/validate.ps1
-> ./eng/test-m0010-tier3.ps1
-> prepare HR-M0010-01 artifacts/evidence
-> obtain actual human review externally
-> ./eng/review-check.ps1 --milestone M0010
-> freshly reread milestone/authority
-> reconcile milestone <-> ledger <-> repository/evidence
-> completion audit
```

Passing automated tests alone is insufficient because the artifact-quality review is blocking.

## Baseline-Executability Audit

Planning confirms:

- M0009 execution ledger records completion and approved real Word review evidence;
- M0009 milestone-index state is reconciled to `complete` by this overlay;
- applicable profiles remain repository-wide `base` + `artifact-first-runtime`;
- repository role remains `product-tool`;
- maturity remains `initial-implementation` until release-readiness work changes it;
- M0010 is an ordinary implementation milestone, not the release-readiness pass;
- heading context and both level-mapping formulas are fixed;
- effective level range/overflow/style-binding behavior is fixed;
- paragraph inheritance behavior is fixed;
- supported value stories/replacement scopes/exclusions are fixed;
- fixture origin/provenance and four-path compatibility matrix are fixed;
- Windows + Word + LibreOffice validation locus is fixed;
- subjective layout acceptance is routed to a blocking human review;
- M0011 release-readiness boundary is fixed;
- remaining choices are local implementation mechanics that can be decomposed into bounded work packages.

The configured baseline model can implement M0010 without deciding new product semantics.

## Escalation Boundary

Return to planning rather than inventing policy if implementation would require changing any of these decisions:

- the relative-heading context definition;
- either `section` or `content` heading-level formula;
- effective heading level limit of 9;
- template heading binding range/role semantics;
- ordinary paragraph inheritance model;
- supported value-story set;
- value replacement logical-scope rule;
- visible-text exclusions;
- adding special Markdown syntax for Word stories;
- application-produced fixture provenance rule;
- required realistic fixture content in a way that weakens story/layout coverage;
- mandatory four-path renderer/origin matrix;
- M0009 Word interop/build approach or renderer/publish contracts;
- validation locus/tier topology;
- blocking human-review policy;
- moving NuGet/release/GitHub workflow work into M0010.

A failure of real Word or LibreOffice to create/finalize the required realistic scenario is evidence for planning, not permission for the implementation agent to simplify the contract silently.
