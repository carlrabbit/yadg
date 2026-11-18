# Box diagrams

YADG converts authored character grids to passive SVG through the existing shared parser. Word uses its figure behavior. PowerPoint places the same vector result as one movable, resizable picture in a template-owned `image` region. Picture components are not individually editable; edit the Markdown source instead.

## Box syntax

Use exact lowercase `diagram`, a stable ID, and an optional caption:

````markdown
```diagram {#service-map caption="Service map"}
┌────────────┐
│ Service    │
│ ┌────────┐ │
│ │ Worker │ │
│ └────────┘ │
└────────────┘
```
````

ASCII is supported:

````markdown
```diagram {#queue}
+--------+
| Queue  |
+--------+
```
````

Whitespace positions every cell. Keep edges aligned. Supported box-drawing characters include `┌ ┐ └ ┘ ├ ┤ ┬ ┴ ┼ ─ │`; connected ASCII forms include `+ - |`. Text can be placed inside or beside boxes.

## Lines, arrows, and overlaps

Connected orthogonal paths and arrowheads such as `→ ← ↑ ↓ ► ◄ ▲ ▼` and ASCII `> < ^ v` are supported. For example:

````markdown
```diagram {#flow}
┌──────┐
│Input │
└──────┘
  ↓
  ↓
```
````

Nested rectangles are drawn as positioned. A missing edge can be reconstructed only when another recognized rectangle fully and unambiguously covers it. Uncovered gaps and ambiguous overlaps are errors; YADG does not guess. The parser tests exercise nested and occluded geometry.

## PowerPoint image regions

Put exactly one diagram in a prepared image region:

````markdown
# Architecture {#text-image}

## Notes {#text}

The grid controls the drawing geometry.

## Diagram {#image}

```diagram {#architecture caption="System architecture"}
┌──────────┐
│ Frontend │
└──────────┘
```
````

One image region accepts one local PNG/JPEG, one identified Mermaid fence, or one qualifying built-in diagram fence. YADG contains the SVG within the template region without changing the aspect ratio or template layout. The caption becomes picture accessibility metadata, not visible slide text. PPTX diagram IDs are deck-scoped and do not join the DOCX global figure/reference namespace. Arbitrary `.svg` files are unsupported.

In DOCX, identified diagrams remain semantic figures with the existing Word template placement, caption, numbering, and reference rules.

## Opt-in auto-detection

Detection defaults off. Enable it in root `YADG.md`:

```yaml
---
yadg:
  version: 1
diagrams:
  autoDetect: true
---
```

When enabled, only unlabeled and exact lowercase `text` fences are candidates, and only strong geometry qualifies. An identified `text {#id caption="..."}` retains its ID; an unlabeled diagram stays anonymous. Weak text such as `A -> B`, prose, or isolated punctuation is not detected. Presentation diagrams are accepted only inside image regions.

## Rejected syntax

An explicit fence without an ID is invalid:

````markdown
```diagram
┌──┐
│ A│
└──┘
```
````

Unknown attributes are also invalid. IDs use `[A-Za-z][A-Za-z0-9_-]*`; `caption="..."` is the only optional attribute. Invalid explicit geometry, duplicate IDs, and packaging errors remain fatal under `--yolo`, with output protected by the authoring transaction.

See [Configuration](CONFIGURATION.md) and [Markdown authoring](MARKDOWN-AUTHORING.md) for the full contracts.
