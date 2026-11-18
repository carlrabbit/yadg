# YADG configuration

`YADG.md` is the canonical workspace marker and configuration filename. It may contain one YAML front matter document followed by optional explanatory text. Schema version remains `1`. Unknown keys, duplicate keys, aliases, anchors, merge keys, custom tags, and malformed YAML are errors.

## Complete workspace schema

```yaml
---
yadg:
  version: 1
values:
  release: "2.4"
markdown:
  thematicBreak: error
  codeInline: style
diagrams:
  autoDetect: false
references:
  headings: number
  figures: number
  tables: label-number
producers:
  mermaid:
    executable: mmdc
    arguments:
      - --quiet
presentations:
  - id: overview
    template: YadgTemplates/overview.pptx
    sources:
      - slides/overview.md
    output: overview.pptx
publish:
  path: ./Published
---
Workspace notes may follow the closing delimiter.
```

All keys are case-sensitive. Values are parsed as YAML and validated by the section that owns them.

| Root | Type and default | Allowed children and behavior |
|---|---|---|
| `yadg` | Mapping; required when front matter exists | `version`: integer `1`. No other child keys. |
| `values` | Optional mapping; empty by default | Stable IDs matching `[A-Za-z][A-Za-z0-9_-]*` map to single-line YAML string scalars. Empty strings are valid. Numbers, booleans, null, sequences, mappings, and multiline values are invalid. Values are literal, nonrecursive data, not secrets. |
| `markdown` | Optional mapping | `thematicBreak`: `error` (default) or `ignore`; `codeInline`: `style` (default), `error`, or `ignore`. Unknown keys or values fail. These settings are not persistent YOLO behavior. |
| `diagrams` | Optional mapping | `autoDetect`: YAML boolean, default `false`. Only unlabeled and exact lowercase `text` fences are candidates; weak candidates remain unsupported fenced code. |
| `references` | Optional mapping | `headings`: `number` (default); `figures`: `number` (default) or `label-number`; `tables`: `number` (default) or `label-number`. Controls generated `[@id]` display and relies on real template caption/numbering mechanics. |
| `producers` | Optional mapping | `mermaid` only. It has required `executable` string and optional `arguments` sequence of strings. This is trusted executable configuration. YADG does not install producer tools. |
| `presentations` | Optional non-empty sequence; no decks by default | Each deck has exactly `id`, `template`, `sources`, and `output`. `id` is unique; `template` is a PPTX template path; `sources` is a non-empty ordered list of Markdown paths; `output` is a unique `.pptx` basename. Declared sources are excluded from DOCX discovery. See [PowerPoint presentations](PRESENTATIONS.md). |
| `publish` | Optional mapping; no default destination | `path`: publication destination. An explicit `--publish-path` overrides it. Publishing copies finalized DOCX and does not parse authoring sources or run producers. |

`YADG.md` without front matter is valid and means empty values, default Markdown policies, auto-detection off, no producer or presentation declarations, and no publication destination. Front matter opens on the first effective line with exactly `---` and closes on the next exact `---`; an unclosed region is an error. UTF-8 BOM is allowed.

## Paths and trust

Workspace template and source paths follow the workspace presentation rules. Local Markdown image paths resolve relative to the source and must remain lexically inside the effective workspace. `publish.path` names a filesystem destination. Producer executable paths and arguments are trusted local configuration and are not sandboxed. Do not place credentials in values or producer arguments.

`workspace.yaml` is not a supported configuration filename. Bootstrap a workspace with `yadg init`; it creates `YADG.md`, `content.md`, and `YadgTemplates/` without replacing existing owned paths.

## Direct `--config`

Direct `yadg presentation check/build --template ... --source ... --output ...` commands do not discover a workspace or read ambient `YADG.md`. An explicit `--config <file>` is one YAML mapping document without Markdown front matter. It is not a workspace marker. Allowed roots are `yadg`, `markdown`, `values`, `producers`, `diagrams`, and `references`; `presentations` and `publish` are not allowed. `yadg.version` must be `1`. Detection is off unless this file explicitly sets `diagrams.autoDetect: true`.

Direct source and template paths resolve from the caller's working directory. Image paths retain direct-authoring root and source-file constraints. Relative producer executable paths and working directories also use the caller's working directory, not the config file directory. No other workspace source or template is discovered. `--yolo` is invocation-local and does not change configuration.

## Examples

Enable diagram auto-detection in workspace sources:

```yaml
---
yadg:
  version: 1
diagrams:
  autoDetect: true
---
```

An explicit direct-authoring config, saved as `authoring.yaml`:

```yaml
yadg:
  version: 1
diagrams:
  autoDetect: true
values:
  release: "2.4"
```

For diagram syntax, see [Box diagrams](BOX-DIAGRAMS.md). For broader workflows, see [Authoring](AUTHORING.md).
