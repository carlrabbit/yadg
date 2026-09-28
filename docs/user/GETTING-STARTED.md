# Getting started

YADG creates Office-independent DOCX files from Markdown and prepared Word templates. Markdown owns content, `YADG.md` owns workspace values and producer settings, and templates own presentation.

## Requirements

YADG 1.0 is a Windows x64 .NET 10 tool. Install the released tool with:

```powershell
dotnet tool install --global Yadg --version 1.0.0
yadg --version
```

Authoring uses .NET and Open XML. Rendering through Word requires activated desktop Word in an interactive user session. Rendering through LibreOffice requires Writer. Mermaid needs a separately installed trusted producer.

## Start a workspace

```powershell
mkdir MyReport
cd MyReport
yadg init
```

`init` creates `YADG.md`, `content.md`, and an empty `YadgTemplates/`. It never overwrites any of those paths. Add a prepared `.docx` template containing `{{content:introduction}}` and run:

```powershell
yadg check --list
yadg build
yadg render --renderer word
# or: yadg render --renderer libreoffice
yadg publish --publish-path ./Published
```

Publishing copies finalized DOCX files; it does not build or render them.

## Learn more

- [Authoring Markdown and workspace configuration](AUTHORING.md)
- [Templates, styles, and lists](TEMPLATES.md)
- [Troubleshooting](TROUBLESHOOTING.md)

## Exploratory best-effort authoring

Strict commands answer whether the workspace is production-correct:

```powershell
yadg check
yadg build
```

For an explicit best-effort artifact while fixing template/runtime issues, use invocation-local `--yolo`:

```powershell
yadg check --yolo
yadg build --yolo
yadg render --renderer word --yolo
```

YOLO reports each degradation and its fallback. It can choose a compatible template resource, a versioned built-in example, a visible placeholder, or the alternate renderer. Missing values/references remain visibly unresolved. YOLO never guesses duplicate IDs, bypasses path safety, claims an authored PreWord was finalized, or makes publishing fall back to PreWords. Production and CI should normally use strict commands. See [troubleshooting](TROUBLESHOOTING.md) and [template introspection](TEMPLATES.md).
