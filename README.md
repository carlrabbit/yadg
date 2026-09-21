# YADG

YADG is a Windows-first, template-first document authoring tool. Markdown owns maintainable content, `YADG.md` owns workspace values and producer settings, and prepared DOCX templates own presentation and Word-native structures. YADG authors an intermediate DOCX, optionally finalizes it through Word or LibreOffice, and explicitly publishes finalized DOCX files.

## V1 installation and prerequisites

YADG V1 is one framework-dependent .NET tool: package `Yadg`, command `yadg`, version `1.0.0`, for Windows x64 and .NET 10. After publication:

```powershell
dotnet tool install --global Yadg --version 1.0.0
yadg --version
```

For a local release candidate:

```powershell
dotnet tool install --tool-path ./.yadg-tool --add-source ./artifacts/package --version 1.0.0 Yadg
./.yadg-tool/yadg.exe --help
```

Authoring and `check`/`build` use .NET and Open XML. The Word renderer requires activated desktop Microsoft Word in an interactive logged-on Windows session; it is not supported from a service, web host, scheduled SYSTEM task, or other unattended/server context. The LibreOffice renderer requires LibreOffice Writer. Mermaid diagrams require a separately installed trusted producer; YADG does not install or bundle it.

## First workflow

A workspace contains `YADG.md`, Markdown sources, and `YadgTemplates/*.docx`. YADG creates `YadgPreWords/*.docx`; renderers create `YadgWords/*.docx`. LibreOffice also retains its existing `YadgPdfs/*.pdf` side output, but PDF is not a YADG publication artifact.

```powershell
yadg check --workspace .
yadg build --workspace .
yadg render --workspace . --renderer libreoffice
# or: yadg render --workspace . --renderer word
yadg publish --workspace . --publish-path ./Published
```

`check` validates without normal outputs. `build` authors DOCX without opening Word. `render` never implicitly builds. `publish` never builds or renders: it copies every top-level finalized `YadgWords/*.docx` to the explicit destination, preserving basenames and unrelated destination files.

## Commands and options

Every command accepts `--workspace <path>` and otherwise uses the current directory.

- `yadg --version` prints the packaged product version; `yadg --help` lists commands.
- `yadg check` validates workspace configuration, Markdown, templates, values, producers, placements, and references.
- `yadg build` authors prepared DOCX files.
- `yadg render --renderer libreoffice|word` finalizes authored DOCX. LibreOffice is the default. `--renderer-path <path>` is LibreOffice-only.
- `yadg publish --publish-path <path>` copies finalized DOCX. The CLI destination overrides `YADG.md publish.path`; without either, publishing fails.

## Workspace and template vocabulary

`YADG.md` can contain schema-v1 YAML front matter followed by human-facing notes:

```yaml
---
yadg:
  version: 1
values:
  document-version: "1.0"
  owner: "Documentation"
publish:
  path: ./Published
---
Workspace notes are not document source.
```

Prepared templates use controls such as `{{content:id}}`, `{{section:id}}`, `{{table:id}}`, `{{figure:id}}`, and `{{table-rows:id}}`. Markdown supports headings with stable IDs, paragraphs, lists, tables, prepared table rows, figures, captions, numeric references such as `[@figure-id]`, and inline Mermaid figure blocks. Template front matter binds existing Word styles and prototypes; it does not define presentation.

For `section` and `content`, heading levels are relative to the nearest preceding template heading at the placement. A `section` includes its selected root; `content` omits that root. Source level gaps are preserved, and effective Word heading levels 1–9 are supported. Ordinary inserted paragraphs remain separate paragraphs and use the placement paragraph's template-owned body presentation.

Workspace values are literal, non-recursive substitutions such as `{{value:owner}}`. They work in ordinary visible text in the main body and tables, all header/footer variants, footnotes, endnotes, comments, and supported Word text boxes/shapes. A tag may span runs within one paragraph/container. Field instructions, metadata, relationship targets, bookmark names, custom XML, image alternative text, and arbitrary attributes are not value targets.

## Content producers and trust

Mermaid is an external process boundary. Configure a trusted executable in `YADG.md`; YADG passes it an input file and expected PNG output, validates the PNG, and embeds it as an ordinary figure. YADG does not sandbox the process, install package managers, or guarantee network isolation. Treat producer configuration as executable project code and review it before running `check` or `build`.

## Renderers and references

LibreOffice is the default compatibility renderer and produces finalized DOCX plus its existing PDF side output. Microsoft Word is the V1 fidelity target and produces finalized DOCX only. The two renderers may differ in pagination and layout; YADG does not claim byte or visual equivalence. Both can update renderer-owned fields, numbering, indexes, and layout-dependent state according to their application capabilities. Word automation owns and cleans up its interactive application instance.

Figures, tables, captions, headings, and references are authored structurally. Numeric references require a uniquely rendered captioned target or numbered heading. Prepared tables let a DOCX template own header rows, widths, borders, and styles while Markdown supplies row data.

## Publishing, security, and limitations

Publishing is a local filesystem copy of finalized top-level DOCX files. It does not publish PDFs, upload to a DMS, sign documents, or perform remote authentication. Values are workspace data rather than a secret store. External producers execute with the workspace's user permissions. Treat templates and producer configuration as trusted input.

V1 is Windows x64/.NET 10 only. It is not a Linux/macOS distribution, self-contained executable, installer, library package, website, or unattended Word automation service. External Mermaid tooling, desktop Word, and LibreOffice are separate prerequisites. Macro-enabled templates and remote/server-side Word automation are outside the supported contract.

## Troubleshooting

- If `check` or `build` reports a missing template, malformed control, value, producer, heading, or reference, fix the workspace/template and rerun `check`.
- If Mermaid fails, verify the configured executable, pinned arguments, PATH/relative path, and valid PNG output.
- If Word rendering fails, run in an interactive logged-on session with activated desktop Word and ensure no modal prompt is blocking automation.
- If LibreOffice rendering fails, install Writer and use `--renderer-path` to identify `soffice` when it is not on PATH.
- If package creation fails, release builds require full Visual Studio MSBuild with `ResolveComReference`/`tlbimp`; `dotnet pack` is not a substitute for generated Office interop.
- Publishing requires at least one top-level finalized DOCX and an explicit or configured destination; reserved YADG output directories are rejected.

## License and release scope

YADG is licensed under the MIT License. The package declares the SPDX license expression `MIT` and includes the repository `LICENSE` file. External NuGet publication, a GitHub Release/tag, and CI/CD automation are separate authorized release operations.
