# Authoring

## Workspace discovery

`YADG.md` is the workspace marker. YADG discovers Markdown sources recursively, excluding hidden directories and `YadgTemplates/`. DOCX templates are top-level `.docx` files in `YadgTemplates/`. Use `yadg check --list` to see the source, template, and semantic-reference inventory in deterministic order.

Stable IDs use `{#id}` on headings, pipe tables, and figures. IDs are workspace-wide; references use `[@id]`. `{{section:id}}` inserts the selected heading and its body; `{{content:id}}` inserts only the body. Flat one-paragraph ordered and unordered lists, pipe tables, images, prepared table rows, values, and Mermaid figures are supported. Nested lists and ordinary fenced code blocks are not.

## Markdown policies

Workspace schema remains version 1. Both options default to `error`, retaining 1.0 behavior:

```yaml
---
yadg:
  version: 1
markdown:
  thematicBreak: ignore
  codeInline: style
---
```

`thematicBreak` accepts `error` or `ignore`. Ignored thematic breaks produce no content. `codeInline` accepts `error`, `ignore`, or `style`. `ignore` preserves code text as ordinary text. `style` keeps inline code distinct and requires a template `codeInline` character-style role only when selected content contains inline code.

## Values and Mermaid

Values are single-line scalars in `YADG.md`, substituted with `{{value:id}}` in visible template text. Mermaid is an external process. Configure a trusted executable explicitly:

```yaml
producers:
  mermaid:
    executable: mmdc
    arguments: []
```

Then use a fenced `mermaid` block with a stable ID. Producer configuration executes with your user permissions; review it before running YADG.

## Check and build

`yadg check` validates without generating normal outputs. It prints errors and warnings with workspace-relative paths and available Markdown/YAML line and column positions. `check --list` adds the discovered inventory; its output is for people and is not a machine-stable format. Reference counts include sections, tables, and figures.

`yadg build` authors `YadgPreWords/*.docx`. `render` finalizes those files using Word or LibreOffice. `publish` copies finalized DOCX files to the selected destination. See [Templates](TEMPLATES.md) for front matter and presentation roles.

## Best-effort mode

`check`, `build`, and `render` accept `--yolo` for one invocation. Strict mode remains the production default. YOLO can omit a thematic break, keep inline code plain, preserve an unresolved `{{value:id}}` or `[@id]`, show a missing figure/Mermaid placeholder, use a related presentation resource or built-in list numbering, and continue with other templates after an unreadable one. Each recovery prints a `degradation` diagnostic and successful commands report the count. `publish` only publishes finalized `YadgWords/*.docx`.

Malformed configuration, duplicate semantic IDs, unsafe paths/reparse points, corrupt package handling, and no useful output remain fatal. YOLO never silently swallows an exception.
