# Milestone — M0004 References and Word Document Structures

## Execution Profile

| Field | Value |
|---|---|
| Lifecycle state | complete |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna |
| Baseline executor readiness | confirmed |
| Decision preservation | confirmed |
| Execution tractability | confirmed |
| Scope size | medium-large |
| Implementation autonomy | high |
| Documentation sync | deferred |
| Focused validation | template-control/prototype, semantic-reference, bookmark, and OOXML-field tests |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real synthetic Markdown + real DOCX templates containing front matter, native field prototypes, and existing TOC/list fields |
| Validation locus/platform | ordinary local/CI .NET environment; Microsoft Word not required |
| Consumer/release validation | not applicable |
| Human review | none required |

## Goal

Add Word-native numbering and cross-reference structures while preserving YADG's template-first ownership model.

M0004 introduces optional visible template front matter, template-owned caption prototypes, numbered figure/table captions, internal Word bookmarks, and semantic numeric cross-references.

Office-independent `build` creates structurally correct field machinery; a later renderer evaluates fields and refreshes TOC/list structures.

## Target State

When complete:

- M0003 templates without front matter remain valid and keep conventional defaults;
- a template may override semantic style roles via visible YAML front matter;
- a template may bind figure/table caption roles to native Word prototype paragraphs;
- prototypes own labels, sequence identifiers, switches, punctuation, style, and formatting;
- table source metadata supports optional caption text;
- captioned figures/tables become numbered targets only when the template has the matching valid prototype;
- `[@stable-id]` authors type-aware numeric Word references;
- figure/table references target bookmarked cloned `SEQ` fields;
- section references work when the Markdown heading is uniquely rendered and effectively numbered;
- internal bookmark names remain private artifact mechanics;
- template TOC/list-of-figures/list-of-tables fields are preserved;
- built field results are explicitly non-authoritative until rendering;
- validation remains Office-independent.

## Scope

- optional template control region;
- YAML front matter schema v1;
- style-role overrides for headings, lists, generated tables, and plain captions;
- named figure/table caption prototype bindings;
- prototype extraction/removal/validation;
- native `SEQ` caption prototypes;
- table `caption="..."` metadata;
- numbered figure/table captions;
- semantic inline `[@id]`;
- type-aware target resolution;
- per-template target cardinality validation;
- internal bookmarks;
- Word `REF` fields;
- numbered section references where a Markdown heading is uniquely rendered and numbered;
- preservation of template-owned TOC/list fields;
- direct public documentation updates required by M0004.

## Non-goals

- Microsoft Word/Office Interop field evaluation;
- `render` or PDF;
- generating TOC/list-of-figures/list-of-tables from scratch;
- binding a Markdown section ID to a separately template-owned heading;
- automatically injecting localized reference labels;
- page-number references;
- deliberate duplicate/clone reference targets;
- prepared-table row population;
- scalar `value` resolution;
- ordinary hyperlinks;
- workspace-global presentation configuration;
- arbitrary prototype types beyond caption prototypes;
- plugins, DMS, or release packaging.

## Decisions and Constraints

- Applicable profiles remain repository-wide `base` and `artifact-first-runtime`; no conflict exists.
- YADG remains a product/tool.
- Labels, styles, numbering conventions, and Word-native presentation are template-owned.
- Front matter binds semantic roles to existing structures; it does not define formatting.
- Front matter is optional; M0003 defaults remain fallback behavior.
- Control content is visible ordinary Word text in a removable main-body prefix.
- Front matter schema version is `1`; unknown keys fail validation.
- Caption prototypes are exactly one Word paragraph containing exactly one `SEQ` field and one logical `{{caption}}` placeholder.
- Prototype literal label, sequence identifier, switches, punctuation, style, and formatting are preserved.
- Figure alt text remains figure caption text.
- Table caption syntax is `{#id caption="text"}`; `{#id}` remains uncaptioned.
- `[@id]` returns only a numeric Word result; labels stay in surrounding prose/template content.
- Figure/table numeric references require non-empty caption, valid configured prototype, and exactly one rendered target.
- Section numeric references require exactly one rendered Markdown heading and effective Word paragraph numbering.
- `content:id` alone does not create a section-number target.
- Binding to separately template-owned headings is deferred.
- Bookmark names are private generated mechanics.
- `build` authors field structures but does not claim current field results.
- Existing template TOC/list fields are preserved, not refreshed.
- Office Interop remains forbidden from authoring.

## Baseline Executor Readiness

Planning has fixed:

- metadata location and schema;
- compatibility defaults;
- style/prototype ownership;
- caption prototype shape;
- table caption source syntax;
- cross-reference syntax and numeric-only meaning;
- target eligibility/cardinality;
- bookmark ownership;
- field-creation versus field-evaluation responsibility;
- section-reference limitation;
- validation and human-review policy.

Implementation owns concrete YAML library, internal types/functions, bookmark-name algorithm, field construction mechanics, cached placeholder result choice, diagnostic codes, tests, and execution decomposition.

## Execution Tractability

At implementation start create and maintain:

```text
.execution/M0004-references-word-structures.md
```

Map every acceptance obligation to bounded work packages and concrete evidence.

## Required Authority

Read:

- `docs/SPECS.md`
- `docs/specs/STRUCTURED-CONTENT.md`
- `docs/specs/WORD-REFERENCES.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`

Do not require the external guide repository or planning conversation.

## Acceptance Criteria

### Front matter and compatibility

- Valid M0003 templates without front matter still check/build with conventional style defaults.
- Front matter occupies the defined initial control region and is absent from output.
- YAML schema v1 is parsed from logical Word paragraph text.
- Unknown keys, malformed YAML, duplicate front matter, and unsupported versions fail `check`.
- Configured style overrides are used and validated when their roles are required.

### Prototypes

- Prototype blocks resolve uniquely by template-local ID and are removed from output.
- `figureCaption`/`tableCaption` bindings resolve to valid prototypes.
- Zero/multiple `SEQ` fields or zero/multiple logical `{{caption}}` placeholders fail `check`.
- A valid clone preserves prototype-owned labels, field instruction/switches, punctuation, paragraph/run formatting, and style.
- Semantic caption substitution does not flatten unrelated prototype formatting.

### Captions and numbering

- Figure alt text remains semantic caption text.
- `{#table-id caption="Interfaces"}` creates table caption text; `{#table-id}` remains valid.
- Captioned objects without configured prototypes retain plain-caption behavior and are not numeric targets.
- Captioned objects with valid configured prototypes emit numbered cloned caption structures.
- Uncaptioned figure/table objects cannot be numeric targets.

### Semantic references

- `[@id]` parses in paragraphs, flat list items, and table cells.
- Resolution is case-sensitive/type-aware.
- Generated reference output is numeric only; no localized label is injected.
- Figure/table `REF` fields target bookmarked cloned sequence targets with hyperlink behavior.
- Referenced figure/table targets must render exactly once in that template.
- Referenced sections must have exactly one rendered Markdown heading and effective Word numbering.
- Section reference uses paragraph-number plus hyperlink semantics.
- A `content:id`-only heading target fails section-reference validation.
- Separately template-owned headings are not silently bound to Markdown section IDs.

### Bookmarks and fields

- Generated bookmark names are valid/unique in real DOCX output.
- Every generated `REF` structurally resolves to a generated bookmark.
- Existing unrelated bookmarks/fields remain intact.
- Generated/prototype `SEQ` and generated `REF` structures are inspectable without Word.

### Template indexes/lists and renderer boundary

- Existing TOC/list-of-figures/list-of-tables fields are preserved.
- M0004 neither creates nor refreshes those indexes.
- Cached displayed field results are not used as correctness evidence.
- No Tier 0-3 validation requires Microsoft Word.

### Existing behavior and hygiene

- M0003 placement semantics remain valid.
- A directly placed numbered object uses its final caption location as its unique target.
- Existing paragraph/list/table/figure behavior remains covered.
- Synthetic fixtures are redistribution-safe.
- Public README/examples are updated where required.
- `.execution/M0004-references-word-structures.md` maps all obligations to evidence.

## Validation

### Tier 0

Build the current solution through the repository's live solution/build interface. Expected: successful build with no unreconciled material warnings.

### Tier 1

Focused tests against real DOCX structures for:

- front-matter extraction/YAML schema;
- style-role resolution;
- simple/complex `SEQ` prototype inspection;
- split-run `{{caption}}` replacement;
- semantic-reference resolution;
- bookmark validity/uniqueness;
- generated `REF` structure;
- target-cardinality/numbering failures.

Record exact commands in the execution ledger.

### Tier 2

```powershell
./eng/validate.ps1
```

Must pass without Microsoft Word.

### Tier 3

Use a synthetic workspace with at least:

1. one compatibility template with no front matter;
2. one M0004 template containing:
   - visible front matter;
   - a non-default style binding;
   - figure/table caption prototypes with real Word `SEQ` fields;
   - existing TOC and/or list field structures;
   - content exercising figure/table/section semantic references.

Run real CLI `check` and `build`.

Structural evidence must prove:

- control/prototype regions removed;
- prototype labels/field structures preserved in clones;
- bookmark pairs and `REF` instructions present;
- valid section paragraph-number reference structure;
- existing template index/list fields preserved;
- placement override interacts correctly with target identity;
- no assertion relies on evaluated field result text.

Invalid scenarios include malformed/unknown front matter, missing/invalid prototype, wrong `SEQ`/caption-placeholder counts, uncaptioned numeric reference, multiply rendered referenced target, and unnumbered section target.

### Tier 4

Not applicable.

### Tier 5

Not required. Final Word-calculated numbering/index appearance belongs to the renderer milestone.

## Constrained Runtime

Small synthetic DOCX templates keep validation bounded.

No service, credential, browser, database, or installed Word runtime is required.

## Research

No durable research artifact is required.

Planning verified generic Word platform facts needed for the boundary: bookmarks are reference targets, `SEQ` provides sequence numbering, and `REF` can return bookmark/paragraph numbering. These are readily rediscoverable platform facts.

All project conclusions are promoted into authority.

## Documentation Impact

Planning updates:

- `docs/SPECS.md`
- `docs/specs/WORD-REFERENCES.md`
- `docs/ARCHITECTURE.md`
- `docs/TERMINOLOGY.md`
- `docs/MILESTONES.md`
- this milestone

Implementation updates direct user-facing usage documentation where necessary.

No `.guide-sync/pending/` hint is currently required.

## Human Review

Applicability: none.

## Completion Expectations

Implementation owns execution decomposition, persistent ledger, validation, fresh authority reread, reconciliation, and completion audit.

## Baseline-Executability Audit

Planning confirms architecture/profile scope, compatibility, metadata/prototype semantics, caption/reference semantics, target eligibility, bookmark/field ownership, renderer boundary, validation evidence, and human-review policy are settled.

Remaining choices are implementation mechanics.

M0004 is ready for the configured baseline implementation model.

## Escalation Boundary

Return to planning if implementation requires a new material decision about:

- front-matter location/schema/compatibility;
- style/prototype ownership;
- caption source/prototype structure;
- semantic-reference syntax/meaning;
- target eligibility/cardinality;
- template-owned-heading binding;
- bookmark public identity;
- Word field/evaluation ownership;
- TOC/list ownership;
- architecture/dependency boundaries;
- validation or human-review policy.
