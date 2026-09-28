# YADG

YADG is a Windows-first, template-first document authoring tool. Markdown owns maintainable content, `YADG.md` owns workspace values and producer settings, and prepared DOCX templates own presentation. YADG authors DOCX, optionally finalizes through Word or LibreOffice, and explicitly publishes finalized DOCX files.

## Install

YADG 1.0.0 is a Windows x64 .NET 10 tool:

```powershell
dotnet tool install --global Yadg --version 1.0.0
yadg --help
```

Authoring uses .NET and Open XML. Word rendering requires activated desktop Word in an interactive logged-on session; LibreOffice rendering requires Writer. Mermaid requires a separately installed trusted producer.

## Workflow

```powershell
mkdir MyReport
cd MyReport
yadg init
# Add a prepared DOCX with {{content:introduction}} in YadgTemplates/
yadg check --list
yadg build
yadg render --renderer word
# or: yadg render --renderer libreoffice
yadg publish --publish-path ./Published
```

`init` never overwrites `YADG.md`, `content.md`, or `YadgTemplates/`. `check` validates without producing outputs; `build` creates authored DOCX; `render` finalizes; `publish` copies finalized DOCX files. Publishing does not build or render.

## User guides

- [Getting started](docs/user/GETTING-STARTED.md)
- [Authoring](docs/user/AUTHORING.md)
- [Templates and styles](docs/user/TEMPLATES.md)
- [Troubleshooting](docs/user/TROUBLESHOOTING.md)

YADG is licensed under MIT. Third-party software has its own licenses and terms. External publication, releases, and CI automation are separate operations.

Exploratory authoring can use `yadg check --yolo`, `yadg build --yolo`, and `yadg render --yolo`. Strict mode is the production default. YOLO prints each degradation and never bypasses identity/path safety or treats PreWords as finalized. Run `yadg inspect template` to inspect role resolution and fallback previews.
