# M0009 Microsoft Word Renderer and Publishing — Execution Ledger

## Status

Complete.

## Work packages

1. **Baseline and boundary** — apply the planning overlay, reconcile the required authority, preserve the authoring/renderer boundary, and establish this ledger.
2. **Build-time Word interop and renderer** — add a Windows-only Word renderer project using MSBuild `COMReference`/`ResolveComReference`/`tlbimp`, owned application lifecycle, finalization, staging, diagnostics, and CLI routing.
3. **Workspace publication contract** — extend schema-v1 `publish.path`, add `publish`, enforce path/artifact/failure semantics, and preserve LibreOffice behavior/defaults.
4. **Focused and repository validation** — add integration coverage for routing, Word preflight substitutes, configuration, publish filesystem behavior, and review tooling; run Tier 0–2.
5. **Real Word evidence and closure** — run mandatory interactive Word Tier 3A and real publishing Tier 3B, prepare HR-M0009-01 evidence, obtain external human approval, reread authority, reconcile evidence, and run the completion audit.

## Acceptance evidence matrix

| Acceptance area | Implementation/evidence |
|---|---|
| Renderer selection/platform | `Program.CreateRenderCommand`; `word` routes to `WordRenderer`, default remains `libreoffice`, and `--renderer-path` is rejected for Word. `Yadg.WordRenderer.csproj` is Windows-targeted and uses `COMReference` with `EmbedInteropTypes=false`; Visual Studio MSBuild generated `Interop.Microsoft.Office.Interop.Word.dll` and built the full solution. |
| Word preflight/ownership/finalization | `WordRenderer` creates a new typed `Word.Application`, suppresses visibility/alerts/macros where available, reports `Application.Version`, stages inputs, updates fields/TOCs/lists/headers/footers/pagination, saves DOCX, and closes/quits in `finally`; real evidence is under `artifacts/review/evidence/M0009`. |
| LibreOffice compatibility | Existing `LibreOfficeRenderer` remains unchanged in behavior; M0005 regression was updated only for the newly supported `word` ID; full 42-test validation passes. |
| Workspace publish configuration | `WorkspaceValues.PublishPath` and parser support optional schema-v1 `publish.path`, reject empty/missing/unknown keys, and preserve existing workspaces. `M0009IntegrationTests.Publish_path_schema_accepts_path_and_rejects_unknown_keys`. |
| Publish precedence/path/artifacts/hygiene | `Publisher` implements CLI-over-config precedence, workspace-relative/absolute resolution, reserved-directory rejection, top-level DOCX-only enumeration, staging/replace/preservation, no implicit render/build, and preflight failure hygiene; focused publish tests pass. |
| Documentation/review tooling | `README.md` documents Word prerequisites/invocation and publish precedence. `eng/review-check.ps1` and `eng/review.ps1` recognize M0009 and require Word version/evidence. |
| Human review | Renewed human approval recorded in `.review/records/HR-M0009-01.md` for the expanded artifact hash. |

## Criterion-level evidence

| Criterion | Evidence |
|---|---|
| `word` renderer recognized; LibreOffice default retained | CLI routing in `src/Yadg.Cli/Program.cs`; M0005/M0009 tests. |
| `--renderer-path` invalid for Word | Word branch emits `YADG-RENDER-021`. |
| COM isolated from authoring/Open XML layers | COM references exist only in `src/Yadg.WordRenderer`; `Yadg.Core`/`Yadg.Word` have no Office references. |
| Windows full solution builds | Visual Studio 18 MSBuild build of `Yadg.slnx` passed. |
| Missing Word/runtime diagnostics | `WordRenderer` COM activation catch emits `YADG-WORD-004`; missing input emits `YADG-WORD-003`. |
| Version, ownership, visibility, alerts, macro suppression | `WordRenderer.RenderOnSta`; real test reports Word `16.0`. |
| Bounded execution and cleanup | STA worker join bound, document close, application quit, COM release, temp cleanup. |
| All PreWords finalized; authored inputs untouched | Real Word test and evidence hashes; staged-copy implementation. |
| SEQ/REF/section/index/pagination refresh | `UpdateDocument` updates fields, TOC/list fields, headers/footers, and repaginates; expanded real-Word fixture now includes numbered section, figure/table SEQ, REF, TOC/list fields, table, and figure. |
| No Word PDF / unrelated output preservation / failed temp hygiene | Word renderer commits only staged DOCX to `YadgWords`; no PDF path; stage cleanup and sibling semantics. |
| Schema-v1 publish path | `WorkspaceValuesParser` plus focused config test. |
| CLI/config precedence and path resolution | `Publisher.Publish` and focused override test. |
| Reserved destination rejection | `Publisher.IsReserved`; focused test proves no destination is created. |
| All top-level DOCX, basename preservation, no recursion/PDF/template copy | Focused publish test with nested DOCX and unrelated file. |
| Destination creation, replacement, unrelated preservation | `File.Move(..., overwrite:true)` after destination-local temp copy; expanded focused tests cover replacement and no leftover temp files. |
| No implicit build/render and failure hygiene | Publisher has no renderer reference; preflight occurs before `Directory.CreateDirectory`; temp cleanup is best effort. |
| README/review tooling/human gate | README and `eng/review*.ps1`; the existing record is intentionally not reused for the changed artifact hash. |

## Validation log

- Tier 0 interop capability probe: `dotnet build src/Yadg.WordRenderer/Yadg.WordRenderer.csproj --configuration Release` reached the mandated `COMReference` item but failed with `MSB4803`: `ResolveComReference` is unsupported by .NET Core MSBuild and requires .NET Framework MSBuild.
- Full-framework probe: `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe eng\WordInterop.proj /t:Build /p:Configuration=Release` failed inside the mandated `ResolveComReference` path because `AxImp.exe` was not found and the installed Windows SDK/tooling required by the task is absent. No `tlbimp.exe` or `AxImp.exe` exists on the machine.
- Tier 2 baseline: `./eng/validate.ps1` passed the existing non-Word repository validation, 36/36 tests, 0 failures. This does not substitute for the blocked Word-enabled build or Tier 3A.
- Tier 2 current: `./eng/validate.ps1` passed the Windows Visual Studio MSBuild build and 42/42 tests, 0 failures.
- Tier 3A expanded run: focused real-Word test passes and covers numbered section, figure/table captions, SEQ/REF, TOC/list fields, table, figure, authored-input preservation, finalized reopen, and no newly owned `WINWORD.EXE` process after completion.
- Tier 3B expanded run: focused publication tests pass and cover configured/CLI destinations, external absolute destination, missing destination, missing inputs, replacement, preservation, recursion exclusion, and temp hygiene.
- Document QA: the bundled `render_docx.py` was attempted but cannot locate a bundled LibreOffice executable on this host; no substitute visual-render claim is recorded. The artifact remains available for the required interactive Word review.
- Fresh closure reread: completed after the expanded implementation and renewed approval.
- Human review: renewed approval recorded through `eng/review.ps1`; the evidence hash matches the current finalized DOCX.

## Completion audit

- Planning overlay, milestone, AGENTS instructions, and all Required Authority documents were freshly reread after the expanded implementation and renewed human approval.
- Expanded acceptance evidence is mapped to implementation, focused tests, full validation, real Word evidence, publication evidence, and the hash-matched approved review record.
- `./eng/validate.ps1` passed with 42/42 tests and the Word-enabled full solution built through Visual Studio MSBuild.
- `./eng/review-check.ps1 --milestone M0009` passed after renewed explicit human approval and verified the current artifact hash.
- No unresolved escalation-boundary blocker remains; M0010 release/NuGet/DMS scope remains outside this milestone.

## Blockers and planning returns

The earlier escalation blocker is resolved by the installed Visual Studio MSBuild/COM wrapper tooling. No new material planning decision was introduced.
