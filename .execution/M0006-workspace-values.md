# M0006 Workspace Values Execution Ledger

## Status

Complete. Planning overlay applied from the requested package. No escalation-boundary decision was required.

## Bounded work packages

1. Workspace source: parse optional `YADG.md` front matter and expose validated values from core.
2. Word validation: recognize inline value tags, resolve the separate case-sensitive namespace, and diagnose unsupported locations without changing block-tag rules.
3. Word authoring: perform literal multi-run substitution in body, table, header, and footer content with first-run formatting inheritance.
4. Documentation and evidence: document direct usage, add focused/integration coverage, run Tier 0–3 validation, and reconcile acceptance evidence.

## Evidence log

| Package / obligation | Evidence |
|---|---|
| Workspace front matter: note-only valid, valid v1, body ignored | `Front_matter_is_optional_and_values_keep_a_separate_case_sensitive_namespace` asserts note-only parsing, valid v1 values, semantic source separation, and that the YADG.md note is absent from the model |
| Front matter: missing close, unsupported version, unknown/duplicate keys, unsupported YAML, invalid ID, non-string/multiline values | `Front_matter_rejects_the_declared_invalid_forms` exercises each declared invalid form; `WorkspaceValuesParser` emits `YADG-VALUES-002` through `YADG-VALUES-011` |
| Empty string accepted | M0006 integration workspace parses `empty: ""` and asserts it |
| Case-sensitive separate namespace; semantic `[@id]` boundary | `Front_matter_is_optional_and_values_keep_a_separate_case_sensitive_namespace` asserts `Owner`/`owner`, a value/object spelling collision, and semantic `[@architecture]` resolution; `WorkspaceValues.Values` uses ordinal keys |
| Missing referenced value fails; unused values allowed | `AnalyzeParagraphs` emits `YADG-VALUE-002`; successful integration workspace includes unused `empty` |
| Inline ordinary text, multiple/repeated tags, split runs | `ReplaceValues` scans logical paragraph text and replaces matches in reverse; M0006 DOCX integration test covers mixed text, repeated location use, and split run |
| Formatting inheritance and unrelated content | `ReplaceLogicalText` inserts at first participating `Text` node and clears only later tag fragments; integration asserts Bold inherited and Italic not inherited |
| Literal/non-recursive replacement | Integration uses `literal: "{{value:owner}}"` and asserts the resulting literal remains once |
| Main body, table cell, header, footer | `AnalyzeParagraphs` and `ReplaceValues` enumerate the four locations; integration asserts all four output texts |
| Unsupported Word locations diagnosed; block placement unchanged | `AnalyzeParagraphs` rejects textbox/other non-value locations and retains body-only block checks; existing M0002–M0005 tests pass |
| `check` validates without output; `build` validates before output modification | CLI validates templates before creating/writing output; invalid-front-matter integration test preserves sentinel output |
| Successful output has no unresolved supported tags; renderer-independent | Integration asserts supported tags are gone, while literal replacement proves non-recursion; no renderer is used or required |
| M0002–M0005 compatibility | `./eng/validate.ps1`: 18 passed, 0 failed |
| Public documentation and synthetic hygiene | README Workspace values section; M0006 fixtures are generated in temporary directories with synthetic values |
| Execution ledger evidence | This file; final authority reread and audit recorded below |

## Validation commands

- Tier 0: `dotnet build Yadg.slnx --configuration Release`
- Tier 1: `dotnet test Yadg.slnx --configuration Release --filter M0006`
- Tier 2: `./eng/validate.ps1`
- Tier 3: `dotnet test Yadg.slnx --no-build --configuration Release --filter M0006`

## Actual validation results

- `dotnet build Yadg.slnx --configuration Release` — passed, 0 warnings, 0 errors.
- `dotnet test Yadg.slnx --no-build --configuration Release --filter M0006` — passed, 4 tests, including the real synthetic workspace/DOCX integration test.
- `./eng/validate.ps1` — passed, 18 tests, 0 failures; no renderer required.
- Tier 3 real synthetic workspace/DOCX coverage is exercised by `M0006IntegrationTests` in the focused run and includes body, table, header, footer, split-run formatting, literal replacement, and failed-validation output preservation.

## Completion audit

Completed: milestone and Required Authority were freshly reread after implementation; repository status/diff was inspected; implementation, tests, README, and this ledger reconcile with the acceptance contract. No human review or Tier 4 validation applies.
