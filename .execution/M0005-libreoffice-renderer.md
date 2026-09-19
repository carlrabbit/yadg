# M0005 — LibreOffice Renderer and Finalization

Status: AWAITING HUMAN REVIEW.

## Authority and implementation boundary

The supplied `YADG-M0005-planning-overlay.zip` was applied before implementation. The active milestone and only its Required Authority documents were read: `docs/SPECS.md`, `docs/specs/STRUCTURED-CONTENT.md`, `docs/specs/WORD-REFERENCES.md`, `docs/specs/RENDERING.md`, `docs/ARCHITECTURE.md`, `docs/ENGINEERING.md`, `docs/TERMINOLOGY.md`, and `.review/pending/HR-M0005-01.md`. No external guide or planning conversation is implementation authority.

## Work packages

1. Renderer boundary and CLI — add engine-neutral result/renderer contracts and `render --workspace --renderer --renderer-path` without changing build semantics.
2. LibreOffice session bridge — discover/version-check the runtime, start an isolated local UNO session/profile, refresh fields/indexes, save DOCX and export PDF from the same loaded document, and clean up bounded process resources.
3. Artifact lifecycle and diagnostics — preflight top-level PreWords, stage/atomically publish sibling outputs, preserve unrelated files, and report actionable failures.
4. Review substrate — implement listing/showing/recording commands and milestone-scoped review checking without creating an approval record; retain the blocking pending request.
5. Real evidence and documentation — run a synthetic M0004 workspace through build/render, inspect DOCX/PDF structure and rendered pages, write review evidence/manifest, and update public usage documentation.

## Acceptance mapping

| Obligation | Implementation/evidence |
|---|---|
| CLI defaults, renderer selection/path precedence, top-level input discovery | renderer command and focused CLI tests |
| no implicit build; PreWords unchanged; output conventions/replacement | render integration and staged-output tests |
| isolated profile, local API session, version/provenance, bounded cleanup | LibreOffice bridge and runtime integration evidence |
| M0004 field/index refresh and same-session DOCX/PDF | real LibreOffice Tier 3 synthetic workspace evidence |
| failure preflight and actionable diagnostics | focused failure-path tests |
| portable Tier 2 and dependency boundaries | `./eng/validate.ps1`, project references/static inspection |
| pending human review and review-check gate | `.review/pending/HR-M0005-01.md`, `eng/review.ps1`, `eng/review-check.ps1` |
| review evidence and public docs | `artifacts/review/evidence/M0005/`, README, this ledger |

## Validation record

All implementation work and automated validation are complete. Tier 0–2 remain portable and do not add a LibreOffice or Office Interop dependency to the authoring graph. Evidence:

- `dotnet build Yadg.slnx --configuration Release --no-restore` — passed, 0 warnings, 0 errors.
- `dotnet test Yadg.slnx --configuration Release --no-build` — passed, 14/14.
- `./eng/validate.ps1` — passed, build and 14/14 tests.
- Real Tier 3 focused test `Render_uses_real_libreoffice_to_refresh_fields_and_write_both_artifacts` — passed against `C:\Program Files\LibreOffice\program\soffice.com`.
- Runtime evidence: LibreOffice `26.8.0.3`, Windows `10.0.26200.0`, isolated temporary UNO profile/session per render, finalized DOCX and same-session PDF under `artifacts/review/evidence/M0005/`.
- `artifacts/review/evidence/M0005/review-manifest.md` records input/output provenance and SHA-256 hashes. `artifacts/review/session/M0005-render-fixed/page-1.png` was rendered with the documents verification workflow and visually inspected; the representative image is visibly present.
- `./eng/review-check.ps1 --milestone M0005` intentionally remains blocked because `.review/pending/HR-M0005-01.md` is still pending. No approval or waiver was recorded by the implementation agent.

## Closure reconciliation

The freshly reread M0005 milestone, Required Authority documents, pending review request, implementation, tests, review scripts, evidence manifest, README, and this ledger agree on the LibreOffice renderer boundary, artifact lifecycle, provenance requirements, and blocking human gate. The image defect was traced to malformed `w:p`/`w:drawing` authoring markup, malformed synthetic table structure, and an incorrect relationship-target normalization path; production and fixture authoring now wrap drawings in runs, the fixture is schema-valid before rendering, and the finalized artifact retains and visibly renders the image. Repository evidence covers CLI discovery/path precedence, preflight diagnostics, isolated local UNO rendering, field/index refresh, same-session DOCX/PDF production, cleanup bounds, and review tooling. The remaining completion action is external human review of `HR-M0005-01`; after a durable human decision is recorded, run `./eng/review-check.ps1 --milestone M0005`.
