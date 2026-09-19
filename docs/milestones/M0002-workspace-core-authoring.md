# Milestone — M0002 Workspace and Core Document Authoring

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
| Focused validation | focused workspace/Markdown/OOXML test invocations derived from repository tests |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real synthetic filesystem workspace + real DOCX packages through `yadg check` and `yadg build` |
| Validation locus/platform | ordinary local/CI .NET environment; Microsoft Word not required |
| Consumer/release validation | not applicable; distributable packaging remains unresolved |
| Human review | none required |

## Goal

Turn the M0001 architectural proof into the first workspace-oriented YADG authoring capability.

A user must be able to place multiple Markdown source files and one or more prepared Word templates into a conventional YADG workspace, validate the entire workspace with explicit semantic references, and build structurally correct authored DOCX files without Microsoft Word.

M0002 also replaces M0001's prototype tag/CLI spellings with the explicit product vocabulary and workspace contract now fixed in `docs/SPECS.md`.

## Target State

When M0002 is complete:

- `yadg check` and `yadg build` operate on one workspace selected by `--workspace`, defaulting to the current directory;
- `YADG.md` identifies the workspace root and is not treated as source content;
- Markdown source files are discovered recursively under the project-defined exclusions and merged into one workspace semantic/reference registry;
- stable section IDs are unique across the workspace and case-sensitive;
- the semantic model preserves structured headings, paragraphs, and supported inline semantics rather than flattening selected bodies to strings;
- templates use `{{content:<id>}}` and `{{section:<id>}}`;
- `{{value:<id>}}` is recognized as reserved but unsupported for M0002;
- block tags are validated as standalone logical paragraphs in supported main-document-body locations;
- tag recognition still succeeds when a logical tag is split across multiple OOXML runs;
- selected Markdown headings/paragraphs/inlines are authored as Word block/run structures with the style contract in `docs/SPECS.md`;
- every top-level DOCX template in `YadgTemplates/` produces a same-named authored DOCX in `YadgPreWords/`;
- workspace validation completes before normal build outputs are modified;
- unsupported Markdown and unsupported Word tag locations fail with actionable diagnostics rather than silent flattening/ignoring;
- repository validation remains Office-independent.

## Scope

M0002 covers these product surfaces:

- single-workspace discovery and validation;
- conventional template/output directories;
- recursive multi-file Markdown discovery;
- workspace-wide section/reference registry;
- structured semantic representation for the M0002 Markdown subset;
- explicit `content` and `section` tags;
- reserved-but-unsupported `value` tag recognition;
- block-tag placement validation;
- paragraph, heading, emphasis, strong-emphasis, and line-break Word authoring;
- required Word heading-style validation;
- multi-template workspace build;
- workspace-aware CLI behavior and diagnostics;
- real filesystem + real DOCX integration scenarios;
- direct public documentation updates required by the implemented CLI/tag behavior.

Implementation may refactor the M0001 prototype freely where necessary to satisfy this contract.

## Non-goals

M0002 does **not** implement:

- lists or Word numbering definitions;
- tables;
- images/figures;
- captions;
- Markdown hyperlinks;
- block quotes;
- code blocks;
- raw HTML;
- scalar `value` resolution;
- template control pages/configuration syntax;
- Markdown link/include files;
- cross-workspace sharing;
- parent-directory discovery/batch processing of multiple workspaces;
- tags inside tables, headers, footers, textboxes, footnotes/endnotes, comments, or other non-main-body locations;
- cross-references/bookmarks;
- TOC/list-of-figures rebuilding;
- Microsoft Word/Office Interop rendering;
- PDF generation;
- alternative renderers;
- extension/plugin loading;
- DMS integration;
- final packaging/distribution.

Unsupported constructs or locations that are detectable by the implemented parser/template scanner must produce diagnostics rather than being treated as supported.

## Decisions and Constraints

- The only applicable guide profiles remain repository-wide `base` and `artifact-first-runtime`; no component/surface profile adds obligations for M0002.
- There is no profile conflict requiring project-local resolution.
- YADG remains a product/tool, not a `dotnet-library` product.
- M0001's `{{yadg:section:content:<id>}}` spelling and explicit `--markdown`/`--template`/`--output` CLI are prototype behavior and carry no compatibility obligation into M0002.
- Product tags are case-sensitive and whitespace-free.
- M0002 supports exactly `content` and `section`; `value` is a reserved recognized vocabulary item and fails as unsupported.
- `content`/`section` tags are block tags and must occupy the full logical paragraph aside from whitespace.
- The M0002 supported tag location is a paragraph in the main document body that is not inside a table or other unsupported container.
- Semantic reference identity is workspace-wide and independent of file order and heading display text.
- Source discovery must not traverse symbolic-link/reparse-point directories; candidate source/template links are invalid.
- Nested workspaces are errors rather than recursively composed units.
- The Markdown semantic model must contain structured nodes and must not carry Open XML SDK types.
- Paragraph presentation is inherited from the template insertion anchor's paragraph properties/style.
- Generated headings preserve Markdown levels and use the corresponding existing Word `Heading1` through `Heading6` style IDs; missing required styles are validation errors.
- Emphasis/strong are semantic inline formatting and render as italic/bold runs; hard breaks render as Word line breaks.
- Lists are deliberately excluded because Word numbering/style semantics are a separate structured-content contract, not an implementation detail for M0002.
- `check` and `build` validate the whole workspace; unreferenced source files are not exempt from Markdown-subset validation.
- `build` performs validation before modifying normal output artifacts.
- Office Interop and Microsoft Word remain forbidden dependencies for M0002 authoring/validation.
- Synthetic redistribution-safe fixtures are mandatory.

## Baseline Executor Readiness

The project-level decisions required for implementation are fixed in this milestone and its required authority:

- workspace identity/discovery;
- source/template/output conventions;
- CLI workspace selection;
- tag vocabulary and compatibility boundary;
- reference identity and selection semantics;
- supported Markdown subset;
- block-tag placement;
- Word style/rendering semantics within the Office-independent authoring layer;
- unsupported-feature behavior;
- validation topology and human-review policy.

Implementation-owned choices include:

- concrete types/functions and refactoring shape;
- how the structured semantic node hierarchy is represented internally;
- exact use of Markdig APIs/extensions consistent with the contract;
- diagnostic code allocation for newly introduced error classes;
- test-framework organization;
- temporary-directory/fixture helpers;
- how `check` accumulates independent diagnostics;
- how output staging is implemented before validated transformation;
- focused `eng/` helper additions, if useful.

No stronger implementation model is assumed to resolve an unsettled project decision.

## Execution Tractability

At implementation start, create and maintain:

```text
.execution/M0002-workspace-core-authoring.md
```

The executor derives bounded work packages and maps milestone obligations to implementation and evidence there.

The milestone is one coherent semantic progression from M0001's proof to a usable workspace authoring path. It should not be split merely because the internal refactor touches multiple projects.

## Required Authority

Read these files before implementation:

- `docs/SPECS.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`

M0001's implementation/execution ledger may be inspected as repository history/evidence when useful, but it is not required authority for M0002.

Do not read the external guide repository to recover implementation rules.

## Acceptance Criteria

### Workspace and discovery

- A workspace with root `YADG.md`, at least one discovered Markdown source, and at least one top-level `YadgTemplates/*.docx` can be checked from the workspace root with no explicit path and through `--workspace <path>`.
- `YADG.md`, `YadgTemplates/`, `YadgPreWords/`, and dot-prefixed directories are excluded from Markdown source discovery as specified.
- A nested `YADG.md` below the workspace root causes workspace validation to fail.
- A missing root marker, missing source set, or missing template set fails with actionable diagnostics.
- Source/template symlink or reparse-point inputs are rejected, and discovery does not traverse linked directories, subject to platform capability for validation evidence.

### Semantic model and references

- Multiple Markdown files contribute to one workspace reference registry.
- Duplicate stable IDs across different source files fail validation.
- Stable IDs and tag references are case-sensitive.
- Headings without IDs remain renderable semantic headings when nested inside selected content but cannot be referenced directly.
- `content` selection excludes the referenced heading while preserving all supported nested blocks through the section boundary.
- `section` selection includes the referenced heading and the same section body.
- The semantic model represents headings, paragraphs, and supported inlines structurally rather than through one flattened section-content string.

### Markdown subset

- Valid paragraphs, emphasis, strong emphasis, soft breaks, hard breaks, and ATX headings are semantically preserved and rendered according to the specification.
- The `{#id}` suffix is not emitted as heading display text.
- Lists, tables, images, block quotes, code blocks, raw HTML, and Markdown links cause `check` to fail rather than being flattened or silently discarded.

### Template/tag contract

- `{{content:<id>}}` and `{{section:<id>}}` are recognized with the explicit case-sensitive grammar.
- A logical supported tag split across multiple OOXML runs is recognized correctly.
- M0001's old `{{yadg:section:content:<id>}}` spelling is rejected rather than silently treated as current syntax.
- `{{value:<id>}}` is recognized as reserved but fails with an unsupported-feature diagnostic.
- A block tag mixed with non-whitespace prose in the same paragraph fails validation.
- A YADG-shaped tag in at least one representative unsupported container/location fails validation rather than being ignored.
- An unresolved stable ID fails with a source/template-oriented diagnostic.
- Selected generated headings require the corresponding existing `HeadingN` style; absence of a required style fails `check`.

### DOCX authoring

- Ordinary authored paragraphs use the insertion anchor's paragraph properties/style.
- Generated headings use the template's existing `HeadingN` style corresponding to the Markdown heading level.
- Emphasis and strong emphasis produce italic/bold semantic runs.
- Hard Markdown line breaks produce Word line-break elements.
- Unrelated template paragraphs, styles, and package structure remain intact.
- A successfully authored output contains no unresolved supported `content` or `section` tags.
- No Microsoft Word installation is required.

### Multi-template build and failure semantics

- A valid workspace with at least two top-level templates produces a same-named output for each under `YadgPreWords/`.
- Existing corresponding output files may be replaced on successful build.
- Unrelated files already present in `YadgPreWords/` are not deleted.
- A semantic/template validation failure occurs before normal authored output files are modified.
- An operational I/O failure after writing begins returns failure and an actionable diagnostic; transactional rollback of already written files is not required.

### CLI/public documentation

- `yadg check [--workspace <path>]` and `yadg build [--workspace <path>]` implement the workspace contract.
- Public repository documentation that still advertises the M0001 prototype tag or required single-file CLI arguments is updated as part of M0002 implementation so the repository does not end the milestone with contradictory usage guidance.

### Boundary preservation

- The core semantic model remains free of Open XML SDK element types.
- M0002 authoring projects remain free of Office Interop/Word runtime dependencies.
- All committed scenario content and DOCX fixtures are synthetic and redistribution-safe.
- `.execution/M0002-workspace-core-authoring.md` maps every acceptance obligation to implementation and validation evidence before closure.

## Validation

### Tier 0 — edit sanity

Target: changed .NET projects and repository metadata.

Locus: ordinary local/CI .NET environment.

Invocation:

```powershell
dotnet build Yadg.slnx --configuration Release --no-restore
```

Expected evidence: zero build errors; warnings are resolved or explicitly reconciled when material.

### Tier 1 — focused semantic/workspace/OOXML validation

Targets:

- real Markdig parse trees/selected Markdown semantics;
- real temporary filesystem workspaces;
- real Open XML SDK DOCX packages;
- block-tag reconstruction across split runs;
- structural Word output for headings/paragraphs/inlines.

Locus: ordinary local/CI .NET environment.

Invocation: focused test filters/commands selected by the executor from repository tests; record exact invocations in the execution ledger.

Required evidence covers every focused acceptance area, including invalid scenarios.

### Tier 2 — repository validation

Target: complete Office-independent repository.

Locus: ordinary local/CI .NET environment.

Invocation:

```powershell
./eng/validate.ps1
```

Expected evidence: successful restore/build/test validation with zero exit status.

### Tier 3 — workspace authoring integration

Target: the real CLI operating on a synthetic multi-file workspace containing at least two real DOCX templates.

Locus: ordinary local/CI .NET environment; Microsoft Word is not installed or not invoked.

Representative invocations:

```powershell
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- check --workspace <valid-workspace>
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- build --workspace <valid-workspace>
```

Also execute representative invalid workspaces/templates for duplicate IDs, unsupported Markdown, malformed/unsupported tag placement, missing styles, and validation-before-output behavior.

Expected evidence:

- process exit status;
- diagnostics for invalid scenarios;
- output file set for valid multi-template build;
- structural inspection of authored DOCX packages proving semantic block/run rendering and preservation of unrelated template structure.

A mocked workspace or mocked OOXML object graph is not equivalent Tier 3 evidence for behavior whose correctness depends on the actual filesystem/DOCX representation.

### Tier 4 — consumer/release validation

Not applicable. Packaging/distribution remains unresolved.

### Tier 5 — human review

Not required.

M0002 does not claim final Word layout or visual rendering correctness. Its acceptance contract is semantic and structurally inspectable.

## Constrained Runtime

M0002 validation is expected to remain bounded enough for the repository's ordinary test suite and small synthetic workspace scenarios.

If aggregate validation later grows beyond constrained agent execution, implementation may add bounded focused invocations while preserving `./eng/validate.ps1` as the canonical repository validation. Do not claim aggregate success from partial focused runs.

No external service, credential, installed Microsoft Word runtime, browser, database, or remote environment is required for M0002.

## Research

No durable research artifact is required for M0002.

The material decisions are resolved from current project authority, the implemented M0001 repository state, and explicit product choices made during planning. No external runtime behavior or expensive-to-rediscover experimental evidence is needed to make implementation executable.

If implementation discovers that a required Word/Markdown behavior cannot be established without new evidence and that evidence would change the project contract, return to planning rather than turning implementation experimentation into policy.

## Documentation Impact

Planning updates:

- `docs/SPECS.md` — workspace, explicit tag grammar, Markdown subset, Word block-authoring behavior, and M0002 compatibility boundary;
- `docs/TERMINOLOGY.md` — workspace/block-tag/reference-registry terminology;
- `docs/ENGINEERING.md` — current canonical validation truth and M0002 filesystem validation policy;
- `docs/MILESTONES.md` — M0001 completion and M0002 readiness;
- this milestone.

Implementation must directly update user-facing documentation such as `README.md` when the completed M0002 CLI/tag behavior would otherwise contradict it.

No broad documentation synchronization is required to start or complete M0002.

No `.guide-sync/pending/` hint is created by planning because there is no known deferred documentation obligation outside the direct M0002 documentation impact.

## Human Review

Applicability: none.

Automated semantic, filesystem, and OOXML structural validation can decide M0002 acceptance. Final rendered Word appearance is outside this milestone and therefore is not a human-review gate.

## Completion Expectations

The implementation agent owns milestone closure and will:

```text
execution decomposition
-> persistent execution ledger
-> implement/validate bounded work packages
-> freshly reread M0002 and required authority
-> reconcile milestone <-> ledger <-> repository/evidence
-> completion audit
```

Passing tests alone does not establish completion when an acceptance criterion is unmapped, contradicted, or only partially evidenced.

## Baseline-Executability Audit

Planning confirms:

- architecture and durable component boundaries are already settled;
- workspace semantics and CLI behavior are fixed;
- tag grammar and the M0001 compatibility boundary are fixed;
- Markdown supported/unsupported semantics are fixed;
- stable-reference semantics are fixed;
- Word block/style semantics required by M0002 are fixed;
- profile applicability is known and conflict-free;
- validation targets, loci, commands, fallback boundaries, and evidence are explicit;
- no external runtime/service specialization is required;
- subjective visual acceptance is correctly excluded rather than left to executor judgment;
- no research conclusion remains unpromoted;
- remaining choices are local implementation mechanics;
- the work can be decomposed into bounded packages and resumed through the implementation-owned execution ledger.

M0002 is ready for the configured baseline implementation model.

## Escalation Boundary

Return to planning if implementation requires a new material decision about:

- workspace discovery or configuration semantics;
- public CLI compatibility;
- tag grammar or case/whitespace rules;
- section/content identity or selection boundaries;
- supported Markdown behavior;
- Word style ownership or generated-heading mapping;
- supported Word tag locations;
- output/failure semantics;
- validation target/topology;
- architecture or dependency boundaries.

Implementation owns concrete code/test mechanics, refactoring, diagnostic code allocation, test organization, execution decomposition, and supporting edits that preserve the contract.
