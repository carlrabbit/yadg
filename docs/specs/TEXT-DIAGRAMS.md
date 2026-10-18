# Text Diagrams Specification

## Status

Authoritative for built-in character-grid diagrams introduced by M0014.

## Purpose

YADG supports a compact diagram form whose source remains understandable as plain Markdown text.

The author supplies spatial layout directly. YADG converts recognized character-grid drawing primitives and text into a passive vector diagram; it does not perform graph layout.

## Architectural boundary

Text diagrams are built-in authoring content.

```text
fenced source
-> character grid
-> renderable primitives
-> conservative rectangle/occlusion reconstruction
-> SVG
-> YADG Word authoring
```

This path is independent of external content producers. Mermaid remains governed by `CONTENT-PRODUCERS.md`.

The diagram semantic/geometry model contains no OpenXML types.

## Explicit fenced syntax

````markdown
```diagram {#system-flow caption="System flow"}
┌──────────┐
│ A        │
└────┬─────┘
     ▼
┌──────────┐
│ B        │
└──────────┘
```
````

Rules:

- language token is exact lowercase `diagram`;
- stable ID is mandatory;
- IDs use `[A-Za-z][A-Za-z0-9_-]*`;
- optional `caption` is plain semantic text;
- unknown attributes are errors;
- empty source is an error.

An explicit diagram is a semantic figure. Existing global ID, placement, caption, numbering, bookmark, and `[@id]` reference rules apply.

## Source geometry

The fenced content is a character-cell canvas.

Whitespace is significant.

Supported structural vocabulary includes at least:

```text
─ │ ┌ ┐ └ ┘ ├ ┤ ┬ ┴ ┼
→ ← ↑ ↓ ► ◄ ▲ ▼
```

and connected ASCII compatibility forms:

```text
- | + > < ^ v
```

ASCII punctuation is interpreted structurally only when its connectivity participates in valid geometry. Otherwise it remains text.

Ordinary non-structural characters form text content at their source grid positions.

## Connectivity

Structural cells expose directional connections.

The parser may normalize Unicode and ASCII forms into one directional connectivity representation.

Supported paths are orthogonal.

M0014 does not define diagonal lines or curves.

## Rectangles and containment

A complete axis-aligned closed rectangle is a recognized shape.

A rectangle fully inside another rectangle establishes visual containment.

Containment is visual only; YADG does not infer domain meaning from it.

Containers do not clip child/peer connectors. A connector may cross a rectangle boundary.

## Occlusion and z-order

Source text may express layering by omitting the lower shape's border where a higher rectangle visually covers it.

A missing rectangle edge may be reconstructed only when:

1. visible aligned fragments establish the candidate edge;
2. the gap is geometrically covered by another recognized rectangle;
3. that covering rectangle explains the complete interruption;
4. one unambiguous rectangle/z-order interpretation results.

The covering rectangle is drawn above the reconstructed rectangle.

Do not infer edges through uncovered gaps.

Ambiguous reconstructions are invalid explicit diagrams and are insufficient for auto-detection.

No explicit z-index syntax is defined.

## Rendering product

YADG generates deterministic, finite SVG.

Generated SVG:

- uses vector geometry for lines/boxes/arrows;
- positions text against the source grid;
- has explicit dimensions/viewBox;
- is passive/self-contained;
- contains no scripts;
- contains no external/remote resources;
- requires no network access;
- is ephemeral and is not persisted as workspace source/cache.

M0014 authorizes SVG only for internally generated text diagrams. It does not add general file-backed SVG figure support.

## Workspace auto-detection configuration

Root `YADG.md` schema v1 accepts:

```yaml
diagrams:
  autoDetect: true
```

`diagrams` is optional and accepts exactly one key:

```text
autoDetect
```

The value is a YAML boolean.

Default:

```text
false
```

Unknown keys/non-boolean values are configuration errors.

## Auto-detection candidate scope

When `autoDetect: true`, YADG may inspect only:

- fenced blocks with no language/info string;
- exact lowercase `text` fenced blocks.

A `text` fence may carry the same optional identified-figure attributes:

```text
text {#id caption="..."}
```

No other language-tagged fence participates.

Normal Markdown paragraphs, indented code, tables, lists, and other block kinds are not candidates.

## Auto-detection decision

Auto-detection reuses the real diagram parser and primitive extraction path.

A candidate qualifies only when strong renderable geometry is present, such as:

- a complete/reconstructable rectangle;
- a connected path ending in a supported arrow;
- a branching/joining connected structural component;
- a sufficiently substantial multi-row Unicode box-drawing component.

Numeric thresholds/minimum sizes are implementation-owned.

False negatives are preferred over false positives.

Examples that are insufficient on their own include:

```text
A -> B
```

ordinary key/value records, aligned prose, or isolated punctuation.

If a candidate does not qualify, it retains the ordinary unsupported fenced-block behavior; auto-detection does not turn it into a generic rendered code block.

## Anonymous auto-detected diagrams

An auto-detected fence with no stable ID renders naturally at its source anchor as an anonymous generated visual block.

It has no public semantic identity:

- no global registry entry;
- no direct `{{figure:id}}` placement;
- no `[@id]` reference;
- no caption/numbering/bookmark;
- no `check --list` reference entry.

YADG must not synthesize a public ID from filename/line/ordinal.

A successfully auto-detected `text` fence carrying `{#id ...}` becomes an ordinary semantic figure.

## Figure integration

Identified diagrams follow `STRUCTURED-CONTENT.md` and `WORD-REFERENCES.md`.

They support:

- natural placement;
- one direct `{{figure:id}}` placement;
- direct-placement suppression of natural emission;
- existing caption behavior;
- existing numeric-reference prerequisites.

Anonymous diagrams support natural emission only.

## Validation behavior

`check` validates explicit diagram syntax, geometry, IDs, captions, configuration, placement/reference prerequisites, and generated SVG authorability.

`build` performs equivalent validation before normal output mutation.

Explicit malformed/ambiguous diagram geometry is a source-positioned error.

Auto-detection ambiguity means "do not auto-detect"; the original fenced-block semantics then apply.

## Compatibility

Existing workspaces are unchanged because `autoDetect` defaults to false.

Existing PNG/JPEG figures remain unchanged.

Existing Mermaid fences and producer behavior remain unchanged.

M0014 does not generally enable arbitrary fenced-code rendering.

## Deferred behavior

M0014 does not define:

- graph layout;
- node/edge semantics;
- automatic routing;
- diagonal/curved paths;
- arbitrary free-form shapes;
- per-shape styles/colors;
- user z-index directives;
- arbitrary SVG source files;
- cached generated diagrams;
- external diagram executables.
