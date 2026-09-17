# M0002 Workspace and Core Document Authoring — Execution Ledger

## Status

Complete.

## Work packages

1. **Workspace substrate** — model one `YADG.md` workspace, discover sources/templates with exclusions and path/link checks, and expose workspace-aware CLI defaults/options.
2. **Structured semantic core** — replace flattened M0001 sections with structured headings, paragraphs, and supported inline nodes; parse the M0002 Markdown subset and build a case-sensitive workspace-wide registry.
3. **Template contract and OOXML authoring** — implement explicit `content`/`section` tags, reserved `value` diagnostics, placement/container validation, heading-style checks, and structured paragraph/run/line-break authoring over real DOCX packages.
4. **Multi-template build and failure semantics** — validate the complete workspace before output mutation and build every top-level template to same-named `YadgPreWords` output.
5. **Integration validation and documentation** — exercise real temporary multi-file workspaces, two real templates, CLI check/build and invalid cases; update public usage documentation.
6. **Closure reconciliation** — freshly reread M0002 and required authority, map every acceptance criterion to repository evidence, and perform the completion audit.

## Acceptance mapping

| Acceptance area | Implementation and evidence |
|---|---|
| Single workspace, marker, default/explicit `--workspace` | `WorkspaceLoader`, `Program`, and `M0002IntegrationTests`; real CLI check runs both from workspace CWD and with `--workspace`. |
| Discovery exclusions and recursive multi-file registry | `WorkspaceLoader.IsMarkdownSource`; test workspace includes nested, dot-prefixed, marker, template, output, and two source files. |
| Nested/missing workspace inputs and link policy | `WorkspaceLoader` emits `YADG-WS-001..008`; nested marker test passes. Reparse points are rejected/skipped by attributes; platform link creation was unavailable as a portable test fixture. |
| Duplicate and case-sensitive stable IDs | `MarkdownDocumentParser.Merge` uses ordinal registry and `YADG-REF-002`; invalid build test proves duplicate failure. Case-sensitive lookup is asserted. |
| Structured headings/paragraphs/inlines and section/content semantics | Structured `YadgBlock`/`YadgInline` records, Markdig conversion, `YadgSection.Select`; parser test verifies heading, paragraph, emphasis, strong, hard break, and unsupported list behavior. |
| Explicit tags, obsolete spelling, reserved value, block placement | `WordAuthoring` exact case-sensitive scanners; tests cover split tags, old syntax, mixed prose, table location, and `value` diagnostics. |
| Word style and semantic block/run authoring | `WordAuthoring` validates `HeadingN`, clones anchor paragraph properties, emits heading/paragraph runs, italic/bold properties, and `Break`; real output test inspects the produced package. |
| Multi-template output and preservation/failure ordering | CLI validates all templates before creating output directory; integration test builds two real templates, preserves `keep.txt`, and confirms sentinel output remains after invalid build. |
| Actionable diagnostics and public CLI documentation | Stable diagnostic codes include source/template locations; `README.md` documents workspace CLI and current tags without M0001 prototype arguments. |
| Boundary/hygiene constraints | Core project references only Markdig; Word uses Open XML SDK only; no Office Interop/Word runtime dependency; tests use synthetic temporary workspaces and DOCX packages. |
| Ledger closure | This ledger maps every M0002 acceptance area and records Tier 0–3 evidence below. |

## Validation evidence

- Tier 0: `dotnet build Yadg.slnx --configuration Release` — passed with 0 warnings and 0 errors.
- Tier 1 focused: `dotnet test Yadg.slnx --configuration Release --no-build --filter 'FullyQualifiedName~M0002IntegrationTests'` — passed 7/7.
- Tier 2 repository: `./eng/validate.ps1` — restore/build/test passed with exit 0 and 7/7 tests.
- Tier 3 integration: `M0002IntegrationTests.Real_cli_builds_two_templates_and_preserves_word_structure` launches the real `yadg.exe` against a real temporary multi-file workspace and two real DOCX templates; check/build pass, same-named outputs are inspected for semantic runs, hard breaks, heading style, unresolved tags, and preserved content.
- Invalid integration evidence: duplicate IDs, Markdown links/lists, obsolete/mixed/table tags, reserved `value`, nested workspace marker, and validation-before-output sentinel scenarios are exercised; failures return non-zero and include stable diagnostics.
- Dependency scan: no `Microsoft.Office`, `Office.Interop`, or `Interop.Word` references occur under `src` or `tests`.
- Platform note: source/template reparse-point rejection is implemented using `FileAttributes.ReparsePoint`; the current portable test run did not create a privileged filesystem link fixture.

## Closure reconciliation

After implementation and final validation, the M0002 milestone, `AGENTS.md`, and its required authority files (`docs/SPECS.md`, `docs/ARCHITECTURE.md`, `docs/ENGINEERING.md`, `docs/TERMINOLOGY.md`) were freshly reread. Repository behavior, tests, CLI documentation, and the ledger were reconciled against the milestone. The M0002 non-goals remain outside the implementation: lists/tables/images/links as supported content, values, configuration/include/link semantics, cross-workspace processing, non-body tags beyond diagnostics, rendering/PDF, extensions, DMS, and packaging.

## Blockers and planning returns

None. No material workspace, compatibility, grammar, Markdown, style, output, validation, or architecture decision was changed during implementation.
