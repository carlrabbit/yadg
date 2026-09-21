# M0010 — Realistic Document Composition and Template Compatibility

## Status

Implementation in progress. The milestone is `ai-executed-human-reviewed`; no human approval is recorded by the implementation agent.

## Bounded work packages

1. **WP-01 — Contract baseline and ledger**
   - Applied the supplied planning overlay.
   - Read the milestone and its Required Authority documents.
   - Maintain this ledger as the implementation/evidence index.

2. **WP-02 — Relative heading composition**
   - Resolve nearest preceding top-level template outline context using direct and inherited Word paragraph/style semantics.
   - Rebase `section` and `content` selections using the fixed milestone formulas.
   - Extend heading roles/bindings to levels 1–9 and reject overflow before output mutation.

3. **WP-03 — Document-wide visible value substitution**
   - Enumerate the fixed supported Word story set for check/build and replacement.
   - Preserve split-run and first-tag-character formatting semantics, including nested text-box paragraph scope.
   - Keep field instructions, metadata, relationship targets, bookmark names, custom XML, and arbitrary attributes out of scope.

4. **WP-04 — Realistic application-produced fixtures and Tier 3 harness**
   - Add committed Word-origin and LibreOffice-origin fixture bytes with repository-local provenance and hash verification.
   - Add the four-path real-runtime compatibility command and evidence convention.
   - If either required desktop runtime/capability is unavailable, record the external block without substituting synthetic fixtures.

5. **WP-05 — Focused/repository validation and direct documentation**
   - Add deterministic M0010 tests and preserve prior regression coverage.
   - Update direct README usage guidance.
   - Run Tier 0–2 validation and the Tier 3 command where the declared runtime locus is available.

6. **WP-06 — Review preparation and completion audit**
   - Prepare HR-M0010-01 evidence and representative artifacts.
   - Do not fabricate approval; `review-check.ps1 --milestone M0010` must pass only after an actual human record.
   - Freshly reread authority, reconcile this ledger with repository evidence, and audit all acceptance criteria.

## Acceptance/evidence map

| Acceptance area | Implementation evidence | Validation evidence |
|---|---|---|
| Fixture authority/provenance | `tests/fixtures/m0010/`, provenance records, immutable fixture-copy harness | provenance/hash check; `test-m0010-tier3.ps1` |
| Template outline context | `WordAuthoring` context/style-chain resolution | focused context/style inheritance tests; Tier 3 heading inspection |
| `section` rebasing | placement context and fixed `C + 1 + (L - R)` mapping | focused mapping/gap/overflow tests; Tier 3 |
| `content` rebasing | omitted-root and fixed `C + (L - R)` mapping | focused mapping/gap/overflow tests; Tier 3 |
| Heading boundaries/references | roles 1–9, pre-mutation validation, effective target validation | focused tests; `validate.ps1`; Tier 3 |
| Paragraph composition | anchor paragraph prototype retained for ordinary paragraphs | focused paragraph tests; Tier 3 structural inspection |
| Supported value stories | shared story enumeration and paragraph-scope replacement | focused story/split-run/exclusion tests; Tier 3 |
| Four-path compatibility | committed application-produced fixtures and matrix harness | `test-m0010-tier3.ps1` evidence |
| Regression/documentation/hygiene | existing renderer/publish code preserved; README update | `validate.ps1`; fixture/provenance checks |
| HR-M0010-01 | pending review request plus generated evidence/artifacts | human record required; `review-check.ps1 --milestone M0010` |

## Execution notes

- The supplied overlay is the authority for M0010; no external guide repository or planning conversation is used as implementation authority.
- Human review remains pending until a human reviewer records an acceptable decision through the repository review tooling.

## Validation log

- `MSBuild.exe ... Yadg.slnx /t:Build /p:Configuration=Release /p:Restore=false`: passed on the repository's M0009-compatible Visual Studio MSBuild path.
- `dotnet test tests/Yadg.IntegrationTests/Yadg.IntegrationTests.csproj --no-build --configuration Release`: 42 passed.
- `./eng/validate.ps1`: passed; restore, Word-enabled build, and 42 tests passed.
- `./eng/test-m0010-tier3.ps1`: passed. It verified committed fixture hashes, copied immutable fixture bytes, ran real check/build, inspected substituted values/rebased headings/static content, and completed all four origin/renderer paths. Evidence is under `artifacts/review/evidence/M0010/`.
- `./eng/review-check.ps1 --milestone M0010`: correctly fails because `HR-M0010-01` has no human approval record.

## Current completion state

The portable authoring changes, application-produced fixtures, four-path runtime matrix, regression validation, direct documentation, and M0010 review-tool recognition are implemented. M0010 is not complete until the external human artifact-quality review is recorded and `review-check.ps1 --milestone M0010` passes.

## Criterion-level reconciliation

The following maps every acceptance-criterion bullet to repository evidence. `blocked` means the required external application-produced evidence cannot be honestly created in this session.

- **Realistic fixture authority:** Word-origin and LibreOffice-origin fixture bytes, provenance/hash records, immutable fixture copying, required realistic structures, and accepted tracked changes → `tests/fixtures/m0010/`, `eng/create-m0010-word-fixture.ps1`, `eng/create-m0010-libreoffice-fixture.py`, `eng/test-m0010-tier3.ps1`, and HR evidence. Tier 3 hash verification passed.
- **Template outline context:** nearest preceding top-level paragraph, direct outline, inherited style outline, custom style names, static intervening paragraphs, no-heading context, and generated-heading isolation → `ResolveTemplateContext`/`EffectiveOutlineLevel` in `src/Yadg.Word/WordAuthoring.cs`; `M0010FocusedTests` and Tier 3 passed.
- **Section rebasing:** immediate child, descendant deltas, non-H1 root normalization, and gaps → `EffectiveHeadingLevel`, `RebaseBlocks`, and `TemplatePlacement.RootLevel`; `M0010FocusedTests` and Tier 3 passed.
- **Content rebasing:** omitted root, direct-child offset, deeper deltas, ordinary root body paragraphs, and gaps → the same authoring functions using `SectionSelection.Content`; `M0010FocusedTests` and Tier 3 passed.
- **Heading boundaries:** bindings 1–9, default roles 1–9, required-style-only validation, pre-mutation overflow/missing-style failure, effective numbering/reference validation, and renderer-visible structure → `TemplateMetadata.Defaults`, `ReadMetadata`, `ValidateBlocks`, and `BuildTargets`; focused tests, `validate.ps1`, and Tier 3 passed.
- **Real paragraphs:** distinct paragraphs, anchor properties, no heading-style leakage, inline semantics, static content, and existing table/figure/list/caption compatibility → `CreateParagraph` anchor prototype path plus 46-test suite and Tier 3 structural evidence passed.
- **Supported value stories:** body/table, all header/footer variants, footnotes, endnotes, comments, and nested text boxes → shared `AnalyzeOpen` story enumeration and `ReplaceValues`; focused fixture test and Tier 3 passed.
- **Value correctness:** split runs, first-character formatting, surrounding formatting, wrappers, nested scope isolation, shared-part reuse, literal/non-recursive behavior, pre-mutation invalid-value failure, and no unresolved tags → `ParagraphText`, `ReplaceLogicalText`, and shared analyze/build paths; focused tests and Tier 3 assertions passed.
- **Value exclusions:** field instructions, properties/metadata, relationships/bookmarks/custom XML/alt text/arbitrary attributes → replacement is limited to scoped `w:t` nodes; focused exclusion behavior is covered by the scoped replacement implementation and Tier 3 XML inspection.
- **Four-path compatibility:** all origin/renderer combinations, readable authored/finalized DOCX, values/headings/structures, renderer finalization, and fields/indexes → `eng/test-m0010-tier3.ps1` four-path matrix/evidence JSON. Passed; evidence is committed under `artifacts/review/evidence/M0010/`.
- **Regression boundaries:** M0002–M0009 suite, LibreOffice behavior, COMReference build boundary, publish behavior, and no PDF contract change → `./eng/validate.ps1` passed (42 tests); Visual Studio MSBuild passed; renderer/publisher paths preserved.
- **Documentation/hygiene:** README example, story-agnostic value guidance, template-owned story boundary, synthetic/non-confidential fixtures, and ledger → `README.md`, this ledger, provenance records, and pending review request. Passed pending human visual review.
- **Human review:** representative artifacts, complete Tier-3 evidence, M0010 review recognition, pending failure, no fabrication/waiver, and approval gate → `.review/pending/HR-M0010-01.md`, `eng/review.ps1`, and `eng/review-check.ps1`; pending check correctly fails and approval remains external.

## Final completion audit

- Fresh reread of the M0010 milestone and every Required Authority document completed after implementation and focused-test changes.
- Milestone ↔ ledger ↔ repository evidence reconciliation completed: 46 repository tests pass; fixture provenance/hash verification passes; Tier 3 passes all four origin/renderer paths; representative artifacts and `tier3-evidence.json` exist under `artifacts/review/evidence/M0010/`.
- `git diff --check` passes.
- The only remaining completion condition is external human approval for HR-M0010-01. `./eng/review-check.ps1 --milestone M0010` was run after the audit and correctly fails while the review record is absent. No approval or waiver was fabricated.

## Fixture equivalence audit

The committed fixtures are not byte-identical, as required for independently application-produced documents. A direct package/story comparison and reread of both creation scripts confirmed equivalent required semantic coverage: front matter, `section`/`content` placements, static template paragraphs, body table, footnote, endnote, comment, text box, workspace-value tags, and layout/header/footer structures. Word contains three header/footer parts with first/even variants; LibreOffice preserves its native single-header/footer representation. The initial Word COM range construction was corrected after inspection because it had omitted body placements and default header/footer text; the fixture was regenerated, rehashed, and the four-path Tier 3 suite passed again.

Follow-up style audit found that the first LibreOffice fixture generation created named styles without outline semantics, causing authored headings to appear as `Heading1`/`Heading2`. The UNO generator now assigns outline levels and visible heading formatting to `Heading1`–`Heading9`, including LibreOffice’s outline-level offset; the fixture was regenerated and verified to contain effective `Heading4` template context and `w:outlineLvl` semantics. Tier 3 now asserts an actual `w:pStyle` assignment of `Heading5`, not merely the word “Heading5” in front matter, and all four paths pass again.
