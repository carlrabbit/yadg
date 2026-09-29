# YADG

YADG is a template-first document authoring tool for maintaining Word documents from source-controlled Markdown.

Markdown owns maintainable content. `YADG.md` owns workspace values, Markdown policy, producer configuration, and optional publication configuration. Prepared DOCX templates own presentation, layout, numbering, fields, and other Word-native structure.

YADG authors DOCX without Microsoft Office, can finalize documents through Microsoft Word or LibreOffice, and can explicitly publish finalized DOCX files to a filesystem destination.

## Install

YADG 1.1.0 is distributed as a framework-dependent .NET 10 tool:

```powershell
dotnet tool install --global Yadg --version 1.1.0
yadg --version
```

The authoritative full-product release target is Windows x64.

Authoring uses .NET and Open XML. Microsoft Word rendering requires activated desktop Word in an interactive logged-on Windows session. LibreOffice rendering requires Writer. Mermaid diagrams require a separately configured trusted Mermaid producer.

The package targets `net10.0`; this does not imply that every renderer is supported on every .NET platform. A Word-render request on a non-Windows runtime must fail actionably rather than crashing.

## The basic model

A workspace looks like this:

```text
MyReport/
├── YADG.md
├── content.md
└── YadgTemplates/
    └── report.docx
```

YADG creates generated directories as needed:

```text
YadgPreWords/   authored intermediate DOCX
YadgWords/      finalized DOCX
YadgPdfs/       LibreOffice PDF output
```

`YadgPreWords`, `YadgWords`, and `YadgPdfs` are YADG-owned result sets. Successful build/render operations replace their current generated contents. Keep unrelated persistent files elsewhere.

## Quick start

Create a workspace:

```powershell
mkdir MyReport
cd MyReport
yadg init
```

`init` creates:

```text
YADG.md
content.md
YadgTemplates/
```

It does not overwrite those paths if they already exist.

The starter `content.md` contains an `introduction` section. Add a prepared DOCX file such as `YadgTemplates/report.docx` containing the visible text tag:

```text
{{content:introduction}}
```

Then run:

```powershell
yadg check --list
yadg build
yadg render --renderer word
yadg publish --publish-path ./Published
```

Or finalize through LibreOffice:

```powershell
yadg render --renderer libreoffice
```

## A small complete authoring example

After `yadg init`, extend `YADG.md`:

```yaml
---
yadg:
  version: 1

values:
  document-version: "1.1"

markdown:
  thematicBreak: ignore
  codeInline: ignore

publish:
  path: ./Published

producers:
  mermaid:
    executable: mmdc
    arguments: []
---
```

`values` are optional source-controlled scalar values. `publish.path` is optional. Configure a Mermaid producer only if the workspace contains Mermaid diagrams.

Now author `content.md`:

```markdown
# Introduction {#introduction}

This document describes the reporting process.

## Architecture {#architecture}

The process has two steps:

1. Author the source.
2. Finalize the document.

See Section [@introduction].

| Component | Purpose |
|---|---|
| YADG | Authors DOCX |
| Word | Finalizes Word fields |
{#component-table}

![System context](images/system-context.png){#system-context}
```

Workspace value syntax in supported visible DOCX template text is:

```text
{{value:document-version}}
```

Semantic references use:

```text
[@stable-id]
```

YADG emits the numeric Word result only when the target has the required numbering/reference structure. Write the surrounding label yourself, for example:

```markdown
See Section [@architecture].
```

A prepared Word template may contain:

```text
Document version: {{value:document-version}}

{{section:introduction}}
```

`{{value:document-version}}` is substituted in supported visible template text.

`{{section:introduction}}` inserts the selected heading plus its body. Alternatively:

```text
{{content:introduction}}
```

inserts only the body below that heading.

Template placement and Word presentation remain template-owned.

## Template configuration

YADG template front matter is visible text inside the DOCX, delimited by marker paragraphs:

```text
{{yadg:frontmatter}}
```

then YAML such as:

```yaml
version: 1
styles:
  headings:
    1: "Heading 1"
  lists:
    unordered: "Corporate Bullet"
    ordered: "Corporate Number"
  generatedTable: "Table Grid"
  caption: "Caption"
  codeInline: "Code"
prototypes:
  unorderedListItem: bullet-item
  orderedListItem: number-item
```

and:

```text
{{/yadg:frontmatter}}
```

Style selectors may use an exact internal style ID, Word display name, or alias. Use YADG itself to discover what the template actually contains:

```powershell
yadg inspect styles
yadg inspect template
```

`inspect styles` lists concrete serialized Word styles, names, IDs, aliases, visibility and relevant capabilities.

`inspect template` explains the semantic presentation roles YADG sees: placements, prototypes, resolved presentation sources, human-searchable Word locations, and YOLO fallback candidates.

### Template examples and prototypes

YADG prefers concrete template examples where presentation cannot be reconstructed reliably from metadata.

For example, a list-item prototype is a real bulleted/numbered Word paragraph placed between:

```text
{{yadg:prototype:bullet-item}}
```

and:

```text
{{/yadg:prototype:bullet-item}}
```

The single prototype paragraph contains exactly one logical:

```text
{{item}}
```

and owns its real Word numbering and formatting.

Caption prototypes and prepared-table prototype rows use the same general idea: the template supplies the concrete Word structure YADG should clone or preserve.

See [Templates](docs/user/TEMPLATES.md) for the complete authoring-oriented explanation.

## Commands

### `yadg init`

Bootstrap a new workspace:

```powershell
yadg init
yadg init --workspace C:\work\MyReport
```

It never overwrites existing owned starter paths.

### `yadg check`

Validate without producing normal outputs:

```powershell
yadg check
yadg check --list
yadg check --yolo
```

`--list` prints discovered Markdown sources, top-level DOCX templates, and referenceable semantic objects.

Diagnostics use workspace-relative source paths. Markdown/YAML diagnostics include line/column when available. DOCX diagnostics use recognizable Word context and nearby searchable text where possible.

### `yadg inspect`

Inspect template resources without needing the whole authoring workspace to be valid:

```powershell
yadg inspect styles
yadg inspect styles --template report.docx

yadg inspect template
yadg inspect template --template report.docx
```

Use `inspect styles` when you need concrete Word style IDs/names/aliases.

Use `inspect template` when you need to understand what YADG will actually use for a semantic presentation role.

### `yadg build`

Create authored intermediate DOCX files:

```powershell
yadg build
yadg build --yolo
```

Successful build replaces the current top-level `YadgPreWords/*.docx` generated set.

### `yadg render`

Finalize authored DOCX:

```powershell
yadg render
yadg render --renderer libreoffice
yadg render --renderer libreoffice --renderer-path "C:\Program Files\LibreOffice\program\soffice.com"
yadg render --renderer word
yadg render --renderer word --yolo
```

Default renderer:

```text
libreoffice
```

Both renderers produce finalized DOCX in `YadgWords/`.

LibreOffice additionally produces PDF in `YadgPdfs/`.

A successful render replaces the complete generated finalized result set. A Word render clears stale LibreOffice PDFs rather than leaving them to look current.

### `yadg publish`

Copy finalized DOCX to an explicit filesystem destination:

```powershell
yadg publish --publish-path ./Published
```

or configure:

```yaml
publish:
  path: ./Published
```

and run:

```powershell
yadg publish
```

CLI `--publish-path` wins over configured `publish.path`.

Publishing reads finalized top-level `YadgWords/*.docx`. It does not build, render, parse Markdown, validate templates, or execute Mermaid.

## YOLO mode

Strict mode is the default.

Strict commands answer:

> Is this workspace/template/runtime production-correct?

YOLO is opt-in for `check`, `build`, and `render`:

```powershell
yadg check --yolo
yadg build --yolo
yadg render --renderer word --yolo
```

YOLO answers:

> What is the best useful artifact YADG can truthfully produce from what I have now?

YOLO does not hide errors. A recovered problem is reported as a visible `degradation` together with the fallback YADG selected.

A presentation failure might conceptually look like:

```text
degradation ... requested list presentation "ListBullet" is unavailable;
fallback: template style "Corporate Bullet"
```

If no compatible template list resource exists, YOLO can use conservative built-in real Word numbering:

```text
builtin:unordered-list-v1
builtin:ordered-list-v1
```

Other defined recoveries include:

- preserve inline code as ordinary text;
- omit a thematic break;
- preserve unresolved `{{value:id}}` visibly;
- preserve unresolved semantic reference identity visibly;
- show an obvious placeholder for a missing figure or failed Mermaid producer;
- skip an isolated unusable template when another useful template can still build.

Renderer YOLO can substitute the other supported renderer. For example:

```powershell
yadg render --renderer word --yolo
```

may finish through LibreOffice if Word is unavailable and LibreOffice succeeds.

The command reports requested and actual renderer. It never pretends LibreOffice output was finalized by Word.

YOLO still fails rather than guessing through problems such as duplicate/ambiguous semantic IDs, malformed required configuration, lexically escaping asset paths, corrupt DOCX handling, no useful output, or both renderers failing.

Production/CI checks should normally use strict mode.

## Markdown support

YADG supports the authoring subset documented in [Authoring](docs/user/AUTHORING.md), including:

- headings and paragraphs;
- emphasis/strong text;
- flat ordered/unordered lists;
- pipe tables;
- PNG/JPEG figures;
- semantic references;
- workspace values;
- Mermaid figures through an external producer.

Important current limitations include:

- no nested lists;
- no ordinary fenced-code-block rendering other than Mermaid;
- thematic breaks are either errors or explicitly ignored;
- inline code is either an error, ordinary text, or mapped to a configured character style;
- remote/data-URI figures are unsupported.

## Rendering requirements

### Microsoft Word

Word rendering requires:

- Windows;
- installed and activated desktop Microsoft Word;
- `Word.Application` registered;
- an interactive logged-on user session.

YADG uses late-bound COM and does not ship Office interop DLLs.

It does not claim support for unattended service/server Word automation.

### LibreOffice

LibreOffice rendering requires a compatible Writer installation.

Use `--renderer-path` when `soffice` is not discoverable through the normal runtime path.

## Workspace trust

YADG is a local user-scope tool and treats the selected workspace as trusted input.

Ordinary file links are not rejected merely because they are links. Recursive source discovery does not follow linked child directories, avoiding loops and duplicate traversal.

Explicit figure paths must remain lexically inside the workspace. Remote URLs and data URIs are unsupported.

External producers run with the user's permissions. Treat producer configuration as executable project tooling.

## More documentation

- [Getting started](docs/user/GETTING-STARTED.md)
- [Authoring](docs/user/AUTHORING.md)
- [Templates and presentation](docs/user/TEMPLATES.md)
- [Troubleshooting and YOLO](docs/user/TROUBLESHOOTING.md)
- [Product specifications](docs/SPECS.md)

YADG itself is licensed under MIT. Microsoft Word, LibreOffice, Mermaid tooling, .NET dependencies, and other third-party software remain subject to their own licenses and terms.
