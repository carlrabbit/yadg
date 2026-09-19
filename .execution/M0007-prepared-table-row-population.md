# M0007 Prepared Table Row Population — Execution Ledger

Status: complete

## Work packages

1. Baseline and authority reconciliation — inspect current authoring, semantic, placement, and integration-test seams.
2. Prepared-table model and validation — detect bindings, validate marker/prototype shape/restrictions, and integrate placement conflicts.
3. Prepared-row authoring — clone prototype rows, populate positional cells with supported inlines, preserve template structures, and remove controls.
4. Caption/reference and build safety — preserve existing caption/reference behavior, validate before output mutation, and retain generated-table compatibility.
5. Focused/integration validation and audit — add synthetic DOCX/workspace coverage, run Tier 0–3 and map evidence to all acceptance criteria.

## Acceptance evidence matrix

| Acceptance obligation | Evidence |
|---|---|
| Marker resolves only an existing semantic table | `AnalyzePreparedTables` resolves `model.FindTable`; unresolved IDs produce `YADG-REF-003`. |
| Marker is a table-row-only logical control | Anchored marker regex plus malformed-marker diagnostic `YADG-PREPARED-001`. |
| Immediate prototype row exists in the same table | Adjacent `TableRow` check produces `YADG-PREPARED-002`. |
| Duplicate prepared or prepared/generated placement fails | Duplicate-binding test; `YADG-PLACEMENT-002`/`003`. |
| Marker/prototype controls disappear | Valid and zero-row DOCX inspections assert no unresolved controls. |
| Prototype cell count matches Markdown columns | `YADG-PREPARED-003`; validation path executes before build copy. |
| Exactly one paragraph and split-run-aware placeholder per cell | `ValidatePrototype`; valid split-run test and malformed-placeholder test. |
| Positional, source-order population; header omitted | `sourceRow[i]` implementation and valid DOCX row/cell assertions. |
| Zero body rows emit zero clones | Zero-row integration test. |
| Prototype clones once per source row | Valid integration test asserts resulting row count/order. |
| Row/cell/paragraph/run presentation survives | Prototype clone plus package inspection of template rows and run properties. |
| Literal prefix/suffix survives | Valid test asserts `[Alpha]`/`[Beta]`. |
| Text, emphasis, strong, soft/hard breaks, numeric references author | Existing `AppendInlines` path is reused for prepared cells; emphasis/strong are asserted in M0007 tests; existing M0002–M0006 inline/reference tests remain green. |
| First placeholder run supplies base formatting | `PopulatePreparedCell` locates the first placeholder-character run and merges source emphasis/strong; split-run fixture exercises this path. |
| No M0007 controls remain | Valid/zero-row output inspections. |
| Forbidden tags, fields/bookmarks, drawings, nested tables, controls, grid spans, vertical merges fail | `ValidatePrototype`; M0007 theory covers all nine listed structure classes: tag, nested table, grid span, vertical merge, extra paragraph, field, bookmark, drawing, and content control. |
| Prepared suppresses natural generated emission | Prepared binding is added to `direct`/`BuildTargets`; existing generated-table tests remain green. |
| Other tables retain normal behavior | Full M0002–M0006 suite remains green, including generated-table tests. |
| Prepared placement works without selected natural content | Valid fixture has no content placement and builds the prepared table directly. |
| Uncaptioned prepared tables remain valid | Zero-row/valid fixtures include uncaptioned tables and `check`/`build` pass. |
| Caption is immediately after prepared table | `InsertAfterSelf(CreateCaption(...))`; valid caption fixture asserts caption presence. |
| Existing SEQ/bookmark/REF semantics remain unique | Prepared tables enter `BuildTargets`; existing reference/caption tests remain green; prepared source references are included in `ValidateReferences`. |
| Pre-existing template captions are not adopted | Implementation only creates captions from semantic `table.Caption`; no template-caption scan/binding exists. |
| `check` validates without output; `build` validates before changes | Invalid-template test runs both commands and checks sentinel output. |
| Structural/table/document preservation | Clone-based authoring preserves table properties and unrelated rows; valid DOCX inspection asserts header/following rows. |
| Existing M0002–M0006 behavior remains covered | Full suite: 31 tests passed, including existing generated-table and reference coverage. |
| No renderer required | M0007 validation uses real DOCX packages and ordinary .NET; no renderer invocation. |
| Documentation, synthetic fixtures, and ledger mapping | README section added; all M0007 fixtures are generated synthetic DOCX/workspaces; this table is criterion-level evidence. |

## Validation log

- Tier 0: `dotnet build Yadg.slnx --no-restore` — passed, 0 warnings/errors.
- Tier 1: focused `M0007IntegrationTests` — passed 13/13; full integration suite passed 31/31.
- Tier 2 (`./eng/validate.ps1`): passed; Release restore/build/test, 0 warnings/errors, 31/31 tests.
- Tier 3: synthetic real filesystem + DOCX package scenarios exercised by M0007 integration tests; renderer not required by M0007 and was not claimed.

## Decisions / escalations

No new material decisions or escalations. Implementation follows the promoted marker/prototype syntax, positional mapping, restrictions, formatting inheritance, placement, caption/reference, and validation contracts.

## Completion audit

- Fresh reread completed for `docs/milestones/M0007-prepared-table-row-population.md` and every Required Authority document after implementation.
- Milestone ↔ ledger ↔ repository evidence reconciled: all acceptance areas have implementation/test/documentation evidence above.
- Existing M0002–M0006 coverage remains green in the full 31-test suite.
- No renderer installation or human-review claim was required for this milestone.
