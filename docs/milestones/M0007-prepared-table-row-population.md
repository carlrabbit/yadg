# Milestone — M0007 Prepared Table Row Population

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
| Focused validation | prepared-table marker/prototype analysis, row cloning, cell inline mapping, placement/caption interaction |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real synthetic DOCX templates with prepared tables and malformed prototypes |
| Validation locus/platform | ordinary local/CI .NET environment; no renderer required |
| Consumer/release validation | not applicable |
| Human review | none required |

## Goal

Allow a prepared Word template to own a table's structure/presentation while YADG populates repeated body rows from an existing semantic Markdown table.

## Target State

- `{{table-rows:<id>}}` marker row binds a semantic table to an existing Word table.
- The next row is a prototype; each cell contains one `{{cell}}`.
- Markdown body rows clone/populate the prototype positionally.
- Template table/header/footer rows, widths, borders, styles, and row/cell formatting are preserved.
- Marker/prototype controls disappear.
- Prepared binding suppresses natural generated-table emission and conflicts with any other direct placement of that semantic table.
- Existing semantic caption/reference behavior continues at the prepared table location.
- Existing generated tables remain unchanged.
- Office/rendering is not required.

## Scope

Marker/prototype syntax; positional mapping; zero-row behavior; cloning; existing table-cell inline subset including references; first-placeholder-run base formatting; preservation; placement planning; caption/reference integration; malformed-prototype diagnostics; direct usage documentation.

## Non-goals

Named/header mapping, automatic Word header population, merged prototype cells, multiple prototype variants, nested/subrows, prototype images/drawings, workspace values inside prototype rows, pre-existing template-caption binding, duplicate placement, external data sources, sorting/filtering/grouping/totals, or renderer changes.

## Decisions and Constraints

- Profiles remain repository-wide `base` + `artifact-first-runtime`, guide 0.8.0, baseline GPT-5.6 Luna.
- Existing Markdown pipe table is the only row-data source.
- Markdown header defines semantic column count/order but is not rendered in prepared mode.
- Word template owns visible header rows.
- Marker row contains only `{{table-rows:<id>}}` plus whitespace.
- Immediate next row is the prototype.
- Marker/prototype are removed.
- Prototype physical cell count equals semantic column count.
- Each prototype cell has exactly one ordinary paragraph and one logical `{{cell}}`.
- Mapping is positional.
- Literal prefix/suffix may surround `{{cell}}`.
- Existing source-cell inline subset is supported.
- First placeholder-character run supplies base run properties; Markdown emphasis/strong augments them.
- Prototype row forbids other YADG tags, pre-existing bookmarks/fields, drawings, nested tables, content controls, horizontal grid spans, and vertical merges.
- Prepared binding is a placement override, mutually exclusive with `{{table:<id>}}` or another prepared binding for the same table.
- Natural table emission is suppressed.
- Existing semantic caption is emitted immediately after the prepared Word table and remains the reference target.
- Renderer behavior does not change.

## Execution Tractability

Create and maintain:

```text
.execution/M0007-prepared-table-row-population.md
```

Map every acceptance obligation to bounded work and evidence.

## Required Authority

Read before implementation:

- `docs/SPECS.md`
- `docs/specs/STRUCTURED-CONTENT.md`
- `docs/specs/PREPARED-TABLES.md`
- `docs/specs/WORD-REFERENCES.md`
- `docs/specs/WORKSPACE-VALUES.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`
- `docs/TERMINOLOGY.md`

Do not require the external guide repository or planning conversation.

## Acceptance Criteria

### Binding/prototype
- Marker resolves only an existing semantic table.
- Marker is inside a Word table row and is the row's only non-whitespace logical text.
- Immediate following row in same table exists and is the prototype.
- Duplicate prepared binding, or prepared + `{{table:id}}`, fails.
- Marker/prototype controls are removed.

### Column/data population
- Prototype physical cell count equals Markdown column count.
- Each prototype cell has exactly one paragraph and one logical split-run-aware `{{cell}}`.
- Mapping and row order are positional/source-order.
- Markdown header is not inserted.
- Zero source rows emit zero clones.

### Cell authoring/preservation
- Prototype clones once per semantic row.
- Row/cell/paragraph/run presentation survives.
- Literal prefix/suffix survives.
- Text, emphasis, strong, soft/hard break, and semantic numeric references author correctly.
- First placeholder run provides base formatting.
- No M0007 controls remain after successful build.

### Restrictions
- Other YADG tags, pre-existing bookmarks/fields, drawings, nested tables, content controls, grid spans, or vertical merges in prototype fail validation.

### Placement
- Prepared binding suppresses natural generated-table emission.
- Other tables retain normal placement behavior.
- Prepared placement works even if no selected section/content reaches the natural anchor.

### Captions/references
- Uncaptioned prepared tables remain valid.
- Captioned prepared table emits existing caption immediately after prepared table.
- Existing SEQ/bookmark/REF semantics remain valid and unique.
- Pre-existing template caption is not silently adopted.

### Validation/preservation
- `check` validates without output; `build` validates before changes.
- Invalid prepared template preserves pre-existing outputs.
- Table style/grid/widths/borders and unrelated rows/document structures remain materially intact.
- Existing M0002-M0006 behavior remains covered.
- No renderer installation is required.

### Documentation/hygiene
- README/direct docs explain generated vs prepared table behavior and syntax.
- Fixtures are synthetic.
- Execution ledger maps every criterion to evidence.

## Validation

### Tier 0
Build current solution.

### Tier 1
Focused tests for marker/prototype parsing, split-run controls, column counts, placeholder validation, forbidden structures, cloning, inline rendering, placement conflicts, caption/reference interaction, and preservation.

### Tier 2

```powershell
./eng/validate.ps1
```

### Tier 3

Use a synthetic workspace with:
- a two-column semantic Markdown table with multiple body rows, rich inline cell content, stable ID, and caption;
- a prepared Word table with template-owned header, marker row, prototype row (including split-run `{{cell}}`), distinctive formatting/widths/borders, and a template-owned following row;
- a comparison template/path proving generated-table behavior still works.

Run `check` and `build`, then inspect real DOCX output for row count/order, control removal, data/inlines/reference fields, preservation, caption/SEQ/bookmark structure, absence of duplicate natural table, and continued generated-table behavior.

Negative scenarios: unresolved/type-wrong ID, duplicate placement, missing prototype, column mismatch, missing/duplicate placeholder, extra paragraph, other YADG tag, field/bookmark, drawing, nested table, content control, horizontal/vertical merge.

### Tier 4
Not applicable.

### Tier 5
Not required; claims are deterministic and structurally verifiable.

## Research

No durable research artifact is required; all implementation-affecting decisions are project-local and promoted into authority.

## Documentation Impact

Planning adds/updates `docs/SPECS.md`, `docs/specs/PREPARED-TABLES.md`, `docs/ARCHITECTURE.md`, `docs/TERMINOLOGY.md`, `docs/MILESTONES.md`, and this milestone; it marks implemented M0006 complete.

Implementation updates README/direct usage documentation. No `.guide-sync/pending/` hint is required.

## Human Review

Applicability: none.

## Completion Expectations

Implementation owns execution decomposition, persistent ledger, bounded implementation/validation, fresh authority reread, milestone-ledger-repository reconciliation, and completion audit.

## Baseline-Executability Audit

M0006 completion, architecture/profile scope, source data model, marker/prototype syntax, positional mapping, formatting/preservation, placement interaction, caption/reference behavior, validation, and human-review policy are settled. Remaining decisions are implementation mechanics. M0007 is ready for GPT-5.6 Luna.

## Escalation Boundary

Return to planning for material changes to marker/prototype syntax, column mapping, source/header ownership, prototype restrictions, formatting inheritance, placement semantics, caption/reference behavior, data sources, architecture/dependency boundary, validation policy, or human-review policy.
