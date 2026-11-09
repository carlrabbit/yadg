# PowerPoint presentations

YADG can create editable PowerPoint decks from Markdown and a prepared `.pptx` template. The template owns slide geometry, styling, static artwork and formatting examples. Markdown owns slide order, titles and content.

## Configure named decks

Add a `presentations` sequence to `YADG.md`. Each entry has a unique `id`, a PPTX `template` under `YadgTemplates/`, an ordered non-empty list of Markdown `sources`, and a unique `.pptx` output basename.

```yaml
---
yadg:
  version: 1
presentations:
  - id: overview
    template: YadgTemplates/corporate.pptx
    sources:
      - slides/overview.md
      - slides/results.md
    output: overview.pptx
  - id: technical
    template: YadgTemplates/corporate.pptx
    sources:
      - slides/technical.md
    output: technical.pptx
---
```

Sources and templates may be shared between named decks. Presentation sources are excluded from DOCX source discovery. `yadg check --list` includes the declared deck IDs, templates, ordered source paths and outputs. `yadg check` validates every configured deck; `yadg build` builds every deck and replaces the complete `YadgPresentations/*.pptx` result set after all decks succeed. DOCX, PDF and presentation outputs use separate result sets. `render` and `publish` remain DOCX-only.

## Author slides

Each level-1 heading starts a slide. Its heading text becomes the title. The optional heading ID selects one of the sixteen standard layouts; without one, YADG uses `content`.

```markdown
# Architecture {#text-image}

## Explanation {#text}

### Backend
- Markdig
- Open XML

## Diagram {#image}

![System overview](images/overview.png)
```

Level-2 headings select regions. Their text is only an authoring label and is not shown on the slide. Level-3 and deeper headings are content headings. One-region layouts may use body content directly; multi-region layouts require explicit region headings. Layout and region IDs may repeat across slides.

| Layout | Regions |
|---|---|
| `title`, `content` | `content:text` |
| `section` | none |
| `two-columns` | `left:text`, `right:text` |
| `three-columns` | `left:text`, `center:text`, `right:text` |
| `two-rows` | `top:text`, `bottom:text` |
| `three-rows` | `top:text`, `middle:text`, `bottom:text` |
| `grid-2x2` | `top-left:text`, `top-right:text`, `bottom-left:text`, `bottom-right:text` |
| `image` | `image:image` |
| `text-image`, `image-text`, `image-above-text`, `text-above-image` | `text:text`, `image:image` |
| `table` | `table:table` |
| `text-table` | `text:text`, `table:table` |
| `table-text` | `table:table`, `text:text` |

Text regions support paragraphs, composed bold/italic/strikethrough, inline code, supported external hyperlinks, content headings and nested lists. Image regions accept one local PNG/JPEG Markdown image or one identified Mermaid block. Table regions accept one pipe table with a header and at least one body row; its column count must match the prepared native table. Source values are substituted literally in titles, text, lists and table cells.

## Prepare the template

Create editable prototype slides in PowerPoint or LibreOffice Impress. Each prototype uses one dedicated `{{slide:<layout-id>}}` marker. A title-bearing prototype uses `{{title}}`; each content region uses its plain marker such as `{{left}}`, `{{image}}` or `{{table}}`. Text shapes provide formatting examples for the roles used in Markdown. Image markers are in bounded text shapes. Table markers are the only text in the top-left cell of a native table with a header and body example row.

YADG clones only selected prototypes in Markdown order, removes control markers, preserves other static elements and emits editable text, pictures and tables. Templates need prototypes only for layouts that the source uses. When source content uses inline code with `markdown.codeInline: style`, add a uniquely marked `{{codeInline}}` run to the relevant title/text/table-cell formatting example. Existing templates without inline code do not need this sample.

## Standalone commands

For a one-off deck, use direct commands. They do not read or require `YADG.md` and do not run DOCX authoring or configured producers.

```powershell
yadg presentation check --template ./templates/corporate.pptx --source ./slides/part1.md --source ./slides/part2.md
yadg presentation build --template ./templates/corporate.pptx --source ./slides/part1.md --source ./slides/part2.md --output ./out/deck.pptx
```

Pass one template and one or more ordered sources. `check` does not accept `--output`; `build` requires one. Direct build writes only the named file and preserves an existing output if validation or generation fails.

An optional explicit `--config authoring.yaml` enables schema-v1 authoring settings and values. Direct commands do not read an ambient `YADG.md`. `--yolo` enables only the documented invocation-local recoveries.

For a configured workspace, select one declared deck without checking unrelated deck inputs:

```powershell
yadg presentation check --id overview
yadg presentation build --id overview --yolo
```

Selected build replaces only that deck's output file. Ordinary workspace `yadg build` still replaces the complete configured presentation output set.

Mermaid can replace the single image in an `image` region:

````markdown
## Diagram {#image}

```mermaid {#architecture-flow caption="Architecture flow"}
flowchart LR
  A --> B
```
````

The Mermaid ID must be unique within the deck. Its optional caption is metadata and is not rendered as a slide caption. The trusted configured producer runs during check/build; YADG does not install or sandbox it. Under YOLO, a missing image or unavailable/failed Mermaid producer produces a visible editable placeholder in the image region.

## Diagnostics and limitations

`check` reports missing/unknown layouts or regions, malformed/duplicate markers, incompatible region content, missing assets, and unrepresentable tables with source/slide/region context where available. Fix the Markdown or prototype and rerun the command.

YADG does not measure rendered text, image or table overflow after generation. Existing AutoFit settings are preserved, but YADG does not assert visual fit or automatically shrink content. Open the generated deck in Impress or PowerPoint and review its layout.
