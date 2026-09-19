# Milestone — M0006 Workspace Values

## Execution Profile

| Field | Value |
|---|---|
| Lifecycle state | ready |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna |
| Baseline executor readiness | confirmed |
| Decision preservation | confirmed |
| Execution tractability | confirmed |
| Scope size | medium |
| Implementation autonomy | high |
| Documentation sync | deferred |
| Focused validation | workspace-front-matter parsing, inline tag replacement, formatting/location/error tests |
| Repository validation | `./eng/validate.ps1` |
| Integration validation | real synthetic workspace + real DOCX template with body/table/header/footer value tags |
| Validation locus/platform | ordinary local/CI .NET environment; no renderer required |
| Consumer/release validation | not applicable |
| Human review | none required |

## Goal

Implement the reserved scalar `value` vocabulary using source-controlled workspace values from root `YADG.md`.

M0006 makes common document facts maintainable once per workspace while keeping visual presentation template-owned.

## Target State

- `YADG.md` may contain schema-versioned YAML front matter.
- Existing note-only `YADG.md` remains valid.
- Front matter defines a separate case-sensitive namespace of single-line string values.
- `{{value:<id>}}` is implemented as inline Word-template substitution.
- Value tags may occur among ordinary text and may span OOXML runs.
- Values work in main-body paragraphs, body table cells, headers, and footers.
- Replacement formatting follows the first logical tag-character run.
- Missing values/malformed front matter fail `check`.
- `build` validates before modifying outputs.
- Values are not interpolated into Markdown and have no renderer dependency.
- M0005 remains complete.

## Scope

- root `YADG.md` optional YAML front matter;
- schema version 1;
- `values` mapping;
- value-ID/type validation;
- separate value namespace;
- inline `{{value:<id>}}` recognition across split runs;
- multiple/repeated value tags;
- body/table/header/footer substitution;
- deterministic first-run formatting inheritance;
- literal/non-recursive replacement;
- validation-before-output;
- direct user-facing documentation updates required by the feature.

## Non-goals

M0006 does not implement typed/locale-formatted/multiline/computed values, environment/CLI/template override layers, secret stores, external value sources, Markdown interpolation, text-box/footnote/endnote/comment/property substitution, Word-field generation from scalar values, prepared-table row population, extensions/plugins, publishing/DMS, or a Microsoft Word renderer.

## Decisions and Constraints

- Applicable profiles remain repository-wide `base` and `artifact-first-runtime`; no conflict exists.
- Guide system remains 0.8.0.
- Baseline implementation model remains GPT-5.6 Luna.
- `YADG.md` remains workspace marker + human notes; optional front matter adds workspace metadata without making the body document source.
- Front matter starts only at the beginning of `YADG.md`, uses `---`, and follows `docs/specs/WORKSPACE-VALUES.md`.
- Schema version is 1.
- Unknown/duplicate keys and unsupported aliases/anchors/merge/custom tags are errors.
- Values are case-sensitive YAML string scalars only, single-line; empty string is valid.
- Value IDs use `[A-Za-z][A-Za-z0-9_-]*` in a namespace separate from semantic object IDs.
- `[@id]` never resolves values; `{{value:id}}` never resolves semantic objects.
- Missing referenced values are errors; unused values are allowed.
- Replacement is literal/non-recursive.
- `value` is inline; existing content/section/table/figure tags remain block tags.
- Split-run recognition is required.
- Replacement inherits run properties from the run containing the first logical tag character.
- Supported locations are main body, body table cells, headers, and footers.
- Workspace values define content facts only; templates remain presentation authority.
- Renderer behavior is unaffected.

## Baseline Executor Readiness

Planning has settled value source, schema/versioning, supported YAML subset, ID/type/namespace, missing/unused behavior, inline semantics, Word locations, split-run behavior, formatting inheritance, literal substitution, Markdown boundary, validation, and human-review policy.

Implementation owns concrete YAML library choice, core types, logical-run replacement mechanics, diagnostic codes, file/test organization, and execution decomposition.

## Execution Tractability

At implementation start create and maintain:

```text
.execution/M0006-workspace-values.md
```

Map every milestone obligation to bounded work and concrete evidence.

## Required Authority

Read before implementation:

- `docs/SPECS.md`
- `docs/specs/WORKSPACE-VALUES.md`
- `docs/specs/WORD-REFERENCES.md`
- `docs/ARCHITECTURE.md`
- `docs/ENGINEERING.md`

Do not require the external guide repository or planning conversation.

## Acceptance Criteria

### Workspace front matter

- Note-only `YADG.md` remains valid with empty values.
- Valid front matter parses schema version 1 and values.
- The post-front-matter Markdown body remains ignored as document source.
- Missing closing delimiter, unsupported version, unknown/duplicate key, unsupported YAML feature, invalid ID, non-string value, and multiline value fail validation.
- Empty-string value is accepted.

### Namespace and resolution

- Value IDs are case-sensitive.
- A semantic object ID and value ID with identical spelling may coexist.
- `{{value:id}}` resolves only values; `[@id]` continues to resolve only semantic objects.
- Missing referenced value fails `check`.
- Unused value is allowed.

### Inline Word authoring

- A value tag can occur before/between/after ordinary text.
- Multiple tags in one paragraph and repeated uses resolve.
- A tag split across multiple Word runs resolves.
- Unrelated text/run formatting remains materially intact.
- Replacement uses the run properties of the run containing the first logical tag character.
- Later tag-fragment formatting does not alter replacement formatting.
- Replacement content is literal and not recursively interpreted.

### Supported locations

- Substitution works in main-body paragraph, body table cell, header, and footer.
- Existing block-tag placement restrictions are unchanged.
- Observable tags in explicitly unsupported locations are diagnosed rather than silently treated as supported.

### Validation/build

- `check` validates front matter and all supported-location references without outputs.
- `build` performs equivalent validation before output modification.
- Failed value validation preserves pre-existing normal outputs.
- Successful output has no supported unresolved `{{value:...}}` tags.
- M0002-M0005 behavior remains compatible.
- No LibreOffice or Microsoft Word installation is required.

### Documentation/hygiene

- README/direct usage documentation describes the M0006 `YADG.md` syntax and value tag.
- Fixtures are synthetic/non-confidential.
- The execution ledger maps every acceptance obligation to evidence.

## Validation

### Tier 0 — edit sanity

Build the current solution with existing repository conventions.

### Tier 1 — focused validation

Cover front-matter schema/delimiters, duplicate/unknown/unsupported YAML, value namespace/type, split-run matching, multiple/repeated replacement, formatting inheritance, literal replacement, and unsupported-location diagnostics.

Record exact commands in the execution ledger.

### Tier 2 — repository validation

```powershell
./eng/validate.ps1
```

Must pass without a renderer.

### Tier 3 — real workspace/DOCX integration

Use a synthetic workspace with:

```yaml
---
yadg:
  version: 1
values:
  document-version: "2.3"
  reporting-date: "2026-09-30"
  owner: "Liquidity Risk"
---
```

plus human workspace notes, ordinary Markdown content, and a real DOCX template containing:

- value mixed with body text;
- multiple tags in one paragraph;
- a tag split across differently formatted runs;
- table-cell value;
- header value;
- footer value;
- enough existing M0004-compatible structure to detect regressions.

Representative commands:

```powershell
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- check --workspace <workspace>
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- build --workspace <workspace>
```

Inspect the real DOCX to prove expected values are present, supported tags are gone, locations are correct, unrelated text survives, replacement formatting follows the first-run rule, and unrelated styles/fields/bookmarks/relationships remain materially intact.

Negative scenarios include malformed/unclosed front matter, unsupported version, duplicate key, non-string value, invalid ID, unresolved value, and representative unsupported Word location.

### Tier 4 — consumer/release validation

Not applicable.

### Tier 5 — human review

Not required. M0006 acceptance is deterministic and structural.

## Constrained Runtime

M0006 is Office- and renderer-independent. No external service, credential, browser, database, LibreOffice, or Microsoft Word runtime is required.

## Research

No durable research artifact is required. M0006 is a project semantic/schema decision over existing local YAML/text/OOXML boundaries; all implementation-affecting conclusions are promoted here or into project authority.

## Documentation Impact

Planning updates/adds:

- `docs/SPECS.md`
- `docs/specs/WORKSPACE-VALUES.md`
- `docs/MILESTONES.md`
- this milestone

This also reconciles the milestone index with completed M0005 execution/review evidence.

Implementation updates README/direct public usage documentation.

No `.guide-sync/pending/` hint is required.

## Human Review

Applicability: none. No `.review/` request is created.

## Completion Expectations

Implementation owns execution decomposition, persistent ledger, implementation/validation, fresh authority reread, milestone-ledger-repository evidence reconciliation, and completion audit.

## Baseline-Executability Audit

Planning confirms architecture/profile scope, M0005 state, value source/schema/type/namespace, inline/split-run/formatting semantics, supported Word locations, compatibility/non-goals, validation targets/loci, and human-review policy are settled.

No research conclusion remains unpromoted. Remaining choices are implementation mechanics. M0006 is ready for GPT-5.6 Luna.

## Escalation Boundary

Return to planning if implementation requires a material change to workspace front-matter syntax/schema/versioning, value type/namespace/source, formatting inheritance, supported Word locations, Markdown interpolation, override/secret/external-source behavior, public tag syntax, namespace interaction, architecture/dependency boundaries, validation policy, or human-review policy.

Implementation owns local parser/type/replacement/test mechanics consistent with this contract.
