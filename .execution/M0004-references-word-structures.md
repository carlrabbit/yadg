# M0004 — References and Word Document Structures

Status: Complete.

## Authority and decomposition

The M0004 milestone and its Required Authority were reread before implementation and again for closure. The supplied milestone was `YADG-M0004-planning-overlay.zip`; the repository baseline supplied `docs/ENGINEERING.md` and `docs/specs/STRUCTURED-CONTENT.md`. No external guide or planning conversation was used as implementation authority.

Work packages:

1. Semantic extensions — optional table captions and `[@stable-id]` reference inlines, preserving the OOXML-free model.
2. Template controls — visible front matter extraction, strict schema-v1 diagnostics, template-local style bindings, prototype discovery/validation, and removal before output.
3. Native structures — prototype cloning with split-run caption replacement, template-owned SEQ preservation, generated bookmarks, REF fields, and numbered section targets.
4. Resolution and safety — type-aware target/cardinality validation, placement-aware target identity, output-before-validation protection, and preservation of unrelated template fields.
5. Evidence and documentation — real DOCX integration fixtures and public usage documentation.

## Acceptance reconciliation

| Obligation | Implementation | Evidence |
|---|---|---|
| M0003 compatibility and optional control region | `TemplateMetadata.Defaults`, `ReadMetadata`, control removal | M0004 compatibility integration and full regression suite |
| Schema-v1 YAML, unknown/malformed/duplicate/version diagnostics | strict visible-paragraph parser in `WordAuthoring.cs` | focused invalid-front-matter/prototype test; CLI `check` |
| Style-role overrides and validation | `StyleFor`, style/numbering checks | synthetic `AltHeading`/numbering template build |
| Unique caption prototypes, exact SEQ/placeholder counts | `CaptionPrototype`, prototype scan/diagnostics | split-run real-DOCX fixture and invalid prototype path |
| Prototype-owned labels, switches, punctuation, formatting | clone/adapt path and `ReplaceLogicalText` | output structural assertions for Figure/Table SEQ and caption text |
| Figure/table captions and numbered eligibility | table `Caption`, figure alt text, prototype-gated target map | real figure/table output and reference fields |
| Semantic reference parsing and type-aware resolution | `YadgReference`, global model lookup, `ValidateReferences` | paragraph reference integration and invalid target test |
| Figure/table target cardinality and placement identity | `BuildTargets` with direct-placement suppression | M0003 placement regression plus M0004 reference build |
| Numbered section references | effective style numbering check, heading bookmark, `REF ... \\p \\h` | synthetic numbered-heading DOCX assertions |
| Private valid unique bookmarks and REF linkage | deterministic SHA-256-derived names, bookmark wrapping | real output bookmark-pair and field assertions |
| Existing bookmarks/fields/TOC/list fields preserved | copy-and-transform of template; only YADG controls removed | output TOC field assertion |
| No field evaluation or Word dependency | authoring emits structural fields with cached placeholder only | build/test/validation run on ordinary .NET; no Office references |
| Existing M0003 behavior and hygiene | regression-preserving changes; synthetic fixtures only | complete test suite and repository inspection |
| Documentation and ledger | README M0004 syntax update and this ledger | repository evidence |

## Validation evidence

Commands executed:

- `dotnet build Yadg.slnx --configuration Release --no-restore` — passed, 0 warnings/errors.
- `dotnet test tests/Yadg.IntegrationTests/Yadg.IntegrationTests.csproj --configuration Release --filter FullyQualifiedName~M0004IntegrationTests` — passed, 2/2.
- `dotnet test Yadg.slnx --configuration Release` — passed for the full suite after M0004 fixes.
- `./eng/validate.ps1` — passed; build succeeded with 0 warnings/errors and all 12 integration tests passed.

Tier 3 uses a temporary synthetic workspace with real Markdown, a PNG asset, a compatibility-independent M0004 DOCX template, visible front matter, split-run caption placeholder, real `SEQ` fields, an existing TOC field, style override, figure/table/section references, and real output-package assertions. No assertion uses evaluated field result text.

## Closure audit

Final audit completed: the M0004 milestone and all Required Authority documents were freshly reread; `git diff --check` passed; the ledger, implementation, tests, README, and validation evidence reconcile. M0004 remains Office-independent; Word field evaluation, TOC refresh, PDF, and renderer behavior remain deferred.
