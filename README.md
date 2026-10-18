# YADG

YADG is a template-first document authoring tool. Maintain semantic content in Markdown, configure workspace values and producer behavior in `YADG.md`, and keep layout and Word-native structure in prepared DOCX templates.

YADG authors DOCX without Office, can finalize documents through Microsoft Word or LibreOffice, and explicitly publishes finalized DOCX files to a filesystem destination.

## Install

The YADG 1.2.0 candidate is a framework-dependent .NET 10 tool. After it is available from your NuGet source, install it with:

```powershell
dotnet tool install --global Yadg --version 1.2.0
yadg --version
```

The authoritative full-product validation target is Windows x64. Word rendering requires activated desktop Word in an interactive logged-on session. LibreOffice rendering requires Writer. The package targets `net10.0`; this does not promise every renderer on every .NET platform.

## Start a workspace

```powershell
mkdir MyReport
cd MyReport
yadg init
```

Add a prepared Word template at `YadgTemplates/report.docx` with the visible tag `{{content:introduction}}`, then author Markdown in `content.md`. YADG uses visible text tags rather than Word content controls.

```powershell
yadg check --list
yadg build
yadg render --renderer word
yadg publish --publish-path ./Published
```

LibreOffice is the default renderer and can be selected explicitly with `yadg render --renderer libreoffice`. LibreOffice also produces PDF output. Publishing copies finalized DOCX only; it does not build, render, parse Markdown, or run content producers.

Successful build and render replace their YADG-owned generated result sets (`YadgPreWords`, `YadgWords`, and `YadgPdfs`). Keep unrelated files elsewhere.

## Choose a diagramming path

Built-in text diagrams preserve source-grid geometry and generate internal SVG without an external tool:

````markdown
```diagram {#service-flow}
┌──────────┐
│ Frontend │
└────┬─────┘
     ▼
┌──────────┐
│ Backend  │
└──────────┘
```
````

Mermaid describes diagram semantics and uses a separately configured trusted producer that returns PNG. YADG does not install Mermaid tooling automatically. Both identified forms use ordinary figure placement in the DOCX template; their source and rendering paths differ.

## Strict and YOLO

Strict mode is the default and checks whether the requested operation is production-correct. `check`, `build`, and `render` also accept invocation-local `--yolo` for truthful best-effort output:

```powershell
yadg check --yolo
yadg build --yolo
yadg render --renderer word --yolo
```

YOLO reports each recovery as a visible degradation and still fails when it cannot produce a useful artifact without guessing semantic meaning.

## Guides

- [Getting started](https://github.com/carlrabbit/yadg/blob/main/docs/user/GETTING-STARTED.md) — guided workspace tutorial.
- [Markdown authoring](https://github.com/carlrabbit/yadg/blob/main/docs/user/MARKDOWN-AUTHORING.md) — source grammar, figures, text diagrams, Mermaid, IDs, and references.
- [DOCX template authoring](https://github.com/carlrabbit/yadg/blob/main/docs/user/DOCX-TEMPLATE-AUTHORING.md) — visible tags, front matter, styles, prototypes, and template presentation.
- [Troubleshooting](https://github.com/carlrabbit/yadg/blob/main/docs/user/TROUBLESHOOTING.md) — diagnostics, producer failures, renderers, and YOLO behavior.

The historical [AUTHORING.md](https://github.com/carlrabbit/yadg/blob/main/docs/user/AUTHORING.md) and [TEMPLATES.md](https://github.com/carlrabbit/yadg/blob/main/docs/user/TEMPLATES.md) paths remain as maintained links to the canonical guides.

YADG is licensed under MIT. Microsoft Word, LibreOffice, Mermaid tooling, .NET dependencies, and other third-party software remain subject to their own licenses and terms.
