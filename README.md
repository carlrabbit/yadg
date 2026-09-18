# YADG

YADG is a template-first document authoring tool. It validates and authors a single Office-independent workspace from Markdown sources and prepared DOCX templates.

## Canonical engineering interface

From a PowerShell checkout:

```powershell
./eng/validate.ps1
./eng/check.ps1 --workspace path/to/workspace
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- build --workspace path/to/workspace
```

With no `--workspace`, both commands use the current directory. A workspace contains `YADG.md`, Markdown sources, top-level `YadgTemplates/*.docx`, and generated `YadgPreWords/*.docx` outputs. Templates use `{{content:<stable-id>}}`, `{{section:<stable-id>}}`, `{{table:<stable-id>}}`, and `{{figure:<stable-id>}}`. M0003 supports flat lists, stable-ID pipe tables (`{#table-id}` on the following line), and block PNG/JPEG figures such as `![System context](images/context.png){#context}`. Headings use explicit identity, for example `## Architecture {#architecture}`. The authoring and validation paths do not require Microsoft Word.
