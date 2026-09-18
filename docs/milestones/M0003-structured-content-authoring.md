# Milestone — M0003 Structured Content Authoring

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
| Focused validation | focused structured-Markdown, asset, placement-plan, and OOXML tests |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real synthetic multi-file workspace + real PNG/JPEG assets + real DOCX templates through `yadg check`/`build` |
| Validation locus/platform | ordinary local/CI .NET environment; Microsoft Word not required |
| Consumer/release validation | not applicable; packaging remains unresolved |
| Human review | none required |

## Goal

Extend the usable Office-independent authoring path from paragraphs/headings into structured document content.

M0003 adds flat lists, generated Markdown tables, figures backed by local image assets, semantic figure captions, and deterministic float-like placement in prepared Word templates.

The milestone must preserve the distinction between semantic source position and physical template placement: table/figure tags relocate objects explicitly rather than causing implicit deduplication or copies.

## Target State

When M0003 is complete:

- flat ordered and unordered Markdown lists author through template-owned Word numbering styles;
- pipe tables with explicit stable IDs become structured semantic tables and author as generated Word tables;
- block PNG/JPEG images with explicit stable IDs become semantic figures;
- figure alt text becomes optional semantic caption text;
- the workspace registry enforces one stable-ID namespace across sections, tables, and figures;
- `{{table:<id>}}` and `{{figure:<id>}}` are implemented block tags;
- every table/figure retains a natural Markdown anchor;
- one direct structured-object tag in a template creates a placement override that suppresses that object's natural-anchor rendering for that template;
- duplicate direct placements of one object in one template fail validation;
- asset paths are source-relative, workspace-contained, regular, non-link files;
- image extents follow the deterministic 96-DPI/downscale-only rule in the structured-content specification;
- required `ListBullet`, `ListNumber`, `TableGrid`, and `Caption` template contracts are validated where applicable;
- real generated DOCX packages contain structurally correct lists, tables, images, and captions without Microsoft Word;
- existing M0002 section/content behavior remains compatible apart from its newly supported structured blocks.

## Scope

M0003 covers:

- semantic list/table/figure nodes;
- natural structured-object anchors;
- workspace-wide type-aware reference registry;
- flat ordered/unordered lists;
- pipe-table parsing and cell inline semantics;
- stable table identity syntax;
- generated Word tables using template `TableGrid`;
- block figure parsing with `![caption](path){#id}`;
- PNG/JPEG asset loading and validation;
- deterministic authored image sizing;
- optional figure caption paragraph using template `Caption`;
- `table` and `figure` direct placement tags;
- per-template placement planning before content rendering;
- natural-anchor suppression under an explicit placement override;
- type-correct reference diagnostics;
- validation-before-output behavior for all new inputs;
- direct public documentation updates required by the implemented feature set.

## Non-goals

M0003 does **not** implement:

- prepared Word table row population or a `table-rows` operation;
- table captions;
- figure/table automatic numbering;
- `SEQ` fields;
- bookmarks;
- cross-references;
- list of figures/tables;
- TOC refresh;
- nested lists;
- multi-paragraph list items;
- task lists;
- arbitrary table spans;
- custom table/list style mappings;
- SVG, GIF, TIFF, BMP, remote images, or data URIs;
- inline images;
- arbitrary image resize attributes;
- intentional duplicate/clone rendering of one table or figure;
- automatic free-floating placement chosen by YADG or Word without an explicit template tag;
- scalar `value` resolution;
- prepared-template configuration pages;
- include/link files or cross-workspace sharing;
- Microsoft Word/Office Interop rendering;
- PDF generation;
- DMS integration;
- extension/plugin loading;
- release packaging.

## Decisions and Constraints

- Applicable guide profiles remain repository-wide `base` and `artifact-first-runtime`; no additional scoped profile applies.
- No profile conflict requires resolution.
- YADG remains a product/tool and is not a `dotnet-library` product.
- M0003 follows `docs/specs/STRUCTURED-CONTENT.md` as the authoritative structured-content contract.
- Stable IDs are one case-sensitive workspace-wide namespace across sections, tables, and figures.
- Every M0003 table and figure requires a stable ID.
- Tables and figures are float-like by default in the semantic sense: they have natural source anchors and may be relocated by templates.
- `{{table:id}}` and `{{figure:id}}` are placement overrides, not additional rendering references.
- Placement overrides are computed per template before rendering any `content`/`section` selection.
- Zero placements means normal natural-anchor rendering; exactly one means relocation; more than one is invalid.
- When a placement override exists, every natural-anchor occurrence of that object is suppressed in that template.
- Without a placement override, repeated rendering of a section/content selection may naturally render the object repeatedly; M0003 performs no hidden global deduplication.
- Direct placement is type-correct. A table tag cannot resolve a figure, section, or future scalar value.
- M0003 provides no deliberate clone operation.
- Flat lists use existing effective template numbering contracts through `ListBullet` and `ListNumber`. YADG does not emit literal bullet/number characters as a substitute.
- Generated tables use existing `TableGrid`; YADG does not populate prepared tables.
- Figure captions use existing `Caption` when alt text is non-empty and contain caption text only; numbering/fields are deferred.
- Supported figures are PNG/JPEG local assets only.
- Image targets are resolved relative to their Markdown source file and must remain physically within the workspace.
- Image authored extents use 96-DPI intrinsic dimensions, never upscale, and proportionally downscale to effective text width at the final placement anchor.
- M0003 remains Office-independent.
- The semantic model remains OOXML-free.
- Synthetic redistribution-safe fixtures are mandatory.

## Baseline Executor Readiness

Planning has resolved the material decisions needed for implementation:

- structured-object identity;
- source syntax;
- list scope/numbering ownership;
- table generation/style behavior;
- figure asset/caption/sizing behavior;
- direct tag vocabulary;
- placement-versus-duplication semantics;
- validation ordering;
- supported formats and non-goals;
- human-review policy.

Implementation owns:

- concrete semantic type/class design;
- Markdig integration mechanics;
- OOXML element construction details compatible with the contract;
- effective Word-style/numbering-resolution mechanics;
- image-decoding library choice consistent with supported formats and repository constraints;
- diagnostic code allocation;
- test organization/helpers;
- placement-plan internal data structures;
- bounded execution work packages.

No implementation-affecting policy decision is intentionally deferred.

## Execution Tractability

At implementation start, create and maintain:

```text
.execution/M0003-structured-content-authoring.md
```

Use it to map milestone obligations to bounded work packages, implementation coverage, validation commands, and closure evidence.

M0003 is one coherent expansion of the semantic/OOXML authoring surface and should not be split merely because it spans parser, core model, and Word authoring projects.

## Required Authority

Read before implementation:

- `docs/SPECS.md`
- `docs/specs/STRUCTURED-CONTENT.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`

Do not require the external guide repository or planning conversation.

## Acceptance Criteria

### Stable objects and semantic model

- The semantic model represents flat lists, tables, figures, and natural object anchors structurally without Open XML SDK types.
- Stable IDs remain case-sensitive and globally unique across section/table/figure types.
- A collision between different object types fails validation.
- Table/figure IDs are metadata and are not emitted as document text.

### Lists

- Flat unordered list source authors as list-item paragraphs using `ListBullet`.
- Flat ordered list source authors using `ListNumber`.
- Existing template style/numbering definitions control appearance.
- Missing or invalid effective numbering configuration for a required list style fails `check`.
- Nested, task, and multi-paragraph lists fail as unsupported rather than flattening.

### Tables

- A valid pipe table followed immediately by `{#id}` becomes a referenceable table.
- A table without a stable ID fails `check`.
- Table cells preserve the supported inline subset.
- Generated Word tables use `TableGrid`, preserve header-row semantics, and use automatic table layout.
- Missing `TableGrid` when required fails `check`.
- Table spans/nested block content fail as unsupported.
- No prepared existing Word table is populated or mutated as table data.

### Figures and assets

- `![caption](relative/path.png){#id}` and JPEG equivalent create referenceable figures.
- Alt text becomes optional semantic caption text.
- Empty alt text generates no caption paragraph.
- Markdown image titles, inline images, and images without stable IDs fail or are diagnosed as unsupported according to the specification.
- Asset paths resolve relative to the containing Markdown source.
- Missing, out-of-workspace, linked/reparse, remote/data-URI, or unsupported-format assets fail `check`.
- PNG/JPEG pixel dimensions are read from the real asset.
- Authored OOXML extents follow the 96-DPI, no-upscale, proportional-downscale rule.
- Downscaling uses effective text width at the final placement anchor.
- A non-empty caption uses the existing `Caption` style; missing required style fails `check`.

### Float-like placement

- `{{table:<id>}}` and `{{figure:<id>}}` are recognized across split OOXML runs under the existing block-tag placement contract.
- A direct tag resolves only the matching semantic object type.
- With no direct placement, a selected object's natural anchor renders the object in source order.
- With exactly one direct placement, the object is omitted at natural anchors and rendered at the direct template location.
- The same relocation behavior applies whether the object's containing section is selected through `content` or `section`.
- A direct placement can render an object even when no section/content tag otherwise reaches its natural anchor.
- More than one direct placement for the same object in one template fails `check`.
- The implementation does not silently create a second copy of an object from one direct placement.
- The implementation does not globally deduplicate natural-anchor output when no placement override exists.

### Build/validation semantics

- Placement plans, stable references, asset validation, required styles/numbering, and figure sizing prerequisites are validated before normal outputs are modified.
- A valid workspace with at least two templates can place the same structured object differently in each template because placement is per-template.
- Unrelated template content/styles/package relationships remain intact.
- Successful outputs contain no unresolved supported YADG tags.
- Existing M0002 paragraph/heading/inline behavior remains valid.
- Repository validation requires no Microsoft Word installation.

### Documentation and hygiene

- Public usage documentation is updated to describe the M0003 supported structured-content syntax and tag vocabulary where otherwise contradictory/incomplete.
- Fixtures and image assets are synthetic and redistribution-safe.
- `.execution/M0003-structured-content-authoring.md` maps all acceptance obligations to evidence before closure.

## Validation

### Tier 0 — edit sanity

Target: changed .NET projects and repository metadata.

Locus: ordinary local/CI .NET environment.

Invocation should use the live solution filename. M0002 currently validates through `Yadg.slnx`; the executor must use the repository's canonical current solution rather than restoring an obsolete name.

Expected evidence: successful build and no unreconciled material warnings.

### Tier 1 — focused validation

Targets:

- real Markdig parsing for lists/tables/figures;
- stable registry/type resolution;
- real local PNG/JPEG asset decoding;
- image size calculations;
- effective Word style/numbering resolution;
- placement-plan semantics;
- generated table/image/caption OOXML.

Locus: ordinary local/CI .NET environment.

Invocation: focused test filters chosen from the implemented repository tests and recorded in the execution ledger.

Evidence must cover positive and negative structured-content cases.

### Tier 2 — repository validation

Invocation:

```powershell
./eng/validate.ps1
```

Locus: ordinary local/CI .NET environment.

Expected evidence: canonical restore/build/test flow passes with zero exit status and without Word.

### Tier 3 — real workspace/DOCX integration

Target: real CLI against a synthetic multi-file workspace containing:

- sections/paragraphs from M0002;
- unordered and ordered lists;
- at least one pipe table with stable ID;
- at least one PNG figure with caption;
- at least one JPEG figure or separate focused JPEG scenario;
- at least two real DOCX templates with different placement choices.

Representative commands:

```powershell
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- check --workspace <workspace>
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- build --workspace <workspace>
```

Required structural evidence includes:

- list paragraph style/numbering references;
- generated table structure/style/header semantics;
- embedded image part and drawing extents;
- caption paragraph/style;
- natural-anchor versus direct-placement behavior;
- per-template placement differences;
- absence of duplicate object rendering under one placement override;
- preservation of unrelated template structure.

Invalid scenarios must include representative duplicate placements, type mismatch, ID collision, missing asset, unsupported image format, missing required style/numbering, unsupported nested list, and table without ID.

Mocks are not equivalent evidence for filesystem, image, or DOCX behavior when the real representation is material.

### Tier 4 — consumer/release validation

Not applicable. Packaging remains unresolved.

### Tier 5 — human review

Not required.

M0003 defines deterministic structural authoring rules and does not claim final rendered Word pagination or visual-layout acceptance. Visual finalization remains a later renderer milestone.

## Constrained Runtime

M0003 validation is expected to remain bounded using small synthetic workspaces/assets.

No external service, credential, browser, database, installed Microsoft Word runtime, or remote environment is required.

If repository validation becomes long, focused tests may be run separately for diagnosis, but milestone completion still requires the canonical Tier 2 gate and required Tier 3 evidence.

## Research

No durable research artifact is required.

The M0003 contract is a project design decision over existing Markdown, OOXML, and local-file boundaries already represented in repository engineering policy. No external runtime uncertainty or expensive-to-rediscover evidence is needed to make the milestone executable.

Implementation experiments that would change sizing, placement, numbering ownership, supported formats, or other project contract must return to planning rather than silently becoming policy.

## Documentation Impact

Planning updates:

- `docs/SPECS.md`;
- `docs/specs/STRUCTURED-CONTENT.md`;
- `docs/TERMINOLOGY.md`;
- `docs/MILESTONES.md`;
- this milestone.

Implementation updates `README.md` or other direct user-facing documentation where M0003 behavior would otherwise be undocumented or contradicted.

No separate `.guide-sync/pending/` hint is required by current planning.

## Human Review

Applicability: none.

No review request/record is created.

## Completion Expectations

Implementation owns milestone closure:

```text
execution decomposition
-> persistent execution ledger
-> implement/validate bounded work packages
-> freshly reread M0003 and required authority
-> reconcile milestone <-> ledger <-> repository/evidence
-> completion audit
```

Passing tests alone does not establish completion when acceptance obligations are unmapped or only partially evidenced.

## Baseline-Executability Audit

Planning confirms:

- architecture/component boundaries remain settled;
- profile applicability is known and conflict-free;
- list semantics and numbering ownership are settled;
- table syntax/style/scope are settled;
- figure syntax/assets/captions/sizing are settled;
- float-like placement semantics and uniqueness are settled;
- compatibility with M0002 section/content behavior is explicit;
- validation targets/loci/evidence are explicit;
- no external runtime specialization is required;
- subjective final Word layout is outside milestone acceptance rather than delegated to executor judgment;
- no research conclusion remains unpromoted;
- remaining decisions are implementation mechanics;
- the expected work can be decomposed and resumed through repository-local execution state.

M0003 is ready for the configured baseline implementation model.

## Escalation Boundary

Return to planning if implementation requires a new material decision about:

- stable-object identity or registry scope;
- list nesting/numbering ownership;
- table source syntax, styling, or prepared-table behavior;
- figure syntax, caption source, image formats, sizing, or asset boundaries;
- placement override semantics or duplication behavior;
- public tag vocabulary;
- Word style ownership;
- architecture/dependency boundaries;
- validation target/topology or human-review policy.

Implementation owns local code/test mechanics, type/function structure, parser/OOXML mechanics, diagnostic code allocation, execution decomposition, and supporting edits consistent with the contract.
