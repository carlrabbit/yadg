# YADG

YADG is a template-first document authoring tool. M0001 provides an Office-independent semantic Markdown boundary, visible textual section-content tags, and OOXML authoring over real DOCX packages.

## Canonical engineering interface

From a PowerShell checkout:

```powershell
./eng/validate.ps1
./eng/check.ps1 --markdown path/to/document.md --template path/to/template.docx
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- build --markdown path/to/document.md --template path/to/template.docx --output path/to/authored.docx
```

The supported M0001 visible tag is `{{yadg:section:content:<stable-id>}}`. The corresponding Markdown heading uses explicit identity, for example `## Architecture {#architecture}`. The authoring and validation paths do not require Microsoft Word.
