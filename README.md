# YADG

YADG is a template-first document authoring tool. It validates and authors a single Office-independent workspace from Markdown sources and prepared DOCX templates.

## Mermaid diagrams

Workspaces may configure a trusted external Mermaid producer in `YADG.md`. The command is package-manager neutral; YADG appends its own input/output paths:

```yaml
producers:
  mermaid:
    executable: bun
    arguments:
      - x
      - --bun
      - --package
      - "@mermaid-js/mermaid-cli@11.17.0"
      - mmdc
```

Inline Mermaid blocks become ordinary figures and require a stable ID:

````markdown
```mermaid {#system-flow caption="System flow"}
flowchart LR
    A --> B
```
````

YADG runs the configured executable directly during `check` and `build`; it does not install Bun, npm, or Mermaid. Producer configuration is trusted executable configuration.

## Workspace values

`YADG.md` may begin with schema-versioned YAML front matter. Its body remains human-facing workspace notes and is not document source. M0006 supports case-sensitive, single-line string values:

```yaml
---
yadg:
  version: 1
values:
  document-version: "2.3"
  reporting-date: "2026-09-30"
  owner: "Liquidity Risk"
---
Workspace notes go here.
```

Prepared DOCX templates use values inline with `{{value:document-version}}`. Values can be mixed with ordinary text in body paragraphs, table cells, headers, and footers. They are literal substitutions; values are not interpolated into Markdown.

## Canonical engineering interface

From a PowerShell checkout:

```powershell
./eng/validate.ps1
./eng/check.ps1 --workspace path/to/workspace
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- build --workspace path/to/workspace
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- render --workspace path/to/workspace --renderer libreoffice
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- render --workspace path/to/workspace --renderer word
dotnet run --project src/Yadg.Cli/Yadg.Cli.csproj -- publish --workspace path/to/workspace --publish-path path/to/delivery
```

With no `--workspace`, commands use the current directory. A workspace contains `YADG.md`, Markdown sources, top-level `YadgTemplates/*.docx`, and generated `YadgPreWords/*.docx` outputs. Templates use `{{content:<stable-id>}}`, `{{section:<stable-id>}}`, `{{table:<stable-id>}}`, and `{{figure:<stable-id>}}`. M0003 supports flat lists, stable-ID pipe tables (`{#table-id}` on the following line), and block PNG/JPEG figures such as `![System context](images/context.png){#context}`. M0004 adds table captions (`{#table-id caption="Interfaces"}`), semantic numeric references such as `[@context]`, and optional visible template front matter with template-local caption prototypes. Front matter binds existing Word styles and prototype paragraphs; it does not define presentation. Numeric references require a uniquely rendered captioned target (or a uniquely rendered numbered Markdown heading). YADG authors `SEQ`, bookmarks, and `REF` field structures without evaluating them. The M0005 `render` command discovers LibreOffice from PATH or accepts `--renderer-path`, uses an isolated temporary UNO session, reads top-level `YadgPreWords/*.docx`, and writes `YadgWords/*.docx` plus matching `YadgPdfs/*.pdf` only after a successful render. Rendering never implicitly builds and never requires Microsoft Word or Office Interop. Headings use explicit identity, for example `## Architecture {#architecture}`. The authoring and validation paths do not require Microsoft Word.

## Generated and prepared tables

The Word renderer is the Windows fidelity target and requires an interactive logged-on Windows user session with desktop Microsoft Word installed, activated, and COM-registered. It refreshes top-level `YadgPreWords/*.docx` and writes finalized DOCX only to `YadgWords`; it does not create PDFs. Optional `YADG.md` `publish.path` supplies the publication default, and `--publish-path` overrides it. `publish` copies all top-level `YadgWords/*.docx`, preserves basenames and unrelated destination files, and never builds or renders.

By default, a Markdown pipe table is a generated Word table. To let a DOCX template own the table header, widths, borders, styles, and surrounding rows, put a marker row in an existing Word table and make the immediately following row its prototype:

```text
{{table-rows:interface-matrix}}
```

Each prototype cell must contain one logical `{{cell}}` placeholder. YADG clones that row once for each Markdown body row, maps columns positionally, preserves literal prefix/suffix text and template formatting, and removes the marker/prototype controls. The Markdown header supplies schema only; template rows supply visible headings. Prepared placement suppresses the table's natural/generated emission and can be captioned/referenced using the existing table-caption rules. `{{table:<id>}}` and `{{table-rows:<id>}}` are mutually exclusive for a table in one template.
