# Template Presentation Specification

## Status

Authoritative for YADG 1.1 template-presentation resolution after the M0012 correction.

This supersedes narrower statements that make Word style lookup the primary presentation model.

`TEMPLATE-STYLES.md` remains authoritative for concrete style discovery and exact selector mechanics where it does not conflict here.

## Principle

YADG is template-first.

A template may communicate presentation through:

- placement/context paragraphs;
- explicit YADG prototypes;
- Word styles;
- numbering definitions;
- prepared table structures;
- placeholder run formatting;
- template-owned Word fields/structures.

These are mechanisms beneath one product model:

> Resolve each semantic presentation role from the best applicable template-owned example/resource. Under explicit YOLO mode, continue through deterministic related resources and versioned built-in fallback examples.

Styles are one presentation mechanism, not the product model.

## Terms

### Semantic presentation role

Examples:

```text
body paragraph
heading level 1..9
unordered list item
ordered list item
generated table
prepared table row
figure caption
table caption
inline code
value replacement run
```

### Template example

A concrete template structure cloned or inherited for its presentation/Word mechanics.

Examples include placement paragraphs, caption prototypes, list-item prototypes, prepared-table prototype rows, and placeholder runs.

### Contextual example

Presentation inherited from the concrete insertion location without a separately named prototype.

### Configured resource

A template-owned resource selected through front matter, such as a paragraph/character/table style.

### Built-in fallback example

A conservative, versioned YADG example used only under explicit YOLO when no safe template-owned resource satisfies the role.

### Degradation

A non-fatal YOLO recovery that changes fidelity/presentation/runtime choice while retaining truthful content.

## Resolution

### Strict

```text
explicit template example/prototype
-> contextual template example where applicable
-> configured resource
-> existing role-specific compatibility default
-> error
```

Older role-specific precedence remains where established; e.g. an explicit list-item prototype stays authoritative over list-style binding.

Strict never uses heuristic related-resource selection or built-in YOLO examples.

### YOLO

YOLO first runs the complete strict chain.

If strict fails for a recoverable presentation reason:

```text
compatible related template example/resource
-> versioned built-in fallback example
-> plain/unformatted truthful representation
-> visible placeholder when the semantic object cannot otherwise be represented
-> degradation diagnostic
```

A step is skipped when it would fabricate meaning.

## Determinism

Identical template/workspace/YADG version/mode must yield identical fallback selection.

Only type/capability-compatible candidates qualify.

Capability outranks names.

Within equal capability:

1. stronger role-specific semantic/name hint;
2. normalized visible name;
3. internal style/resource ID.

Do not expose runtime enumeration or hash-map order as fallback behavior.

## Role contracts

### Ordinary paragraph

Preferred source is the concrete placement/insertion paragraph and its paragraph/run properties.

YOLO fallback identity:

```text
builtin:plain-v1
```

meaning ordinary paragraph/plain run presentation without claiming template ownership.

### Headings

Strict:

```text
configured heading role
-> HeadingN compatibility default
-> error
```

YOLO may choose a paragraph style with matching effective outline level.

If none exists, it may preserve heading text in a plain paragraph with the necessary direct outline level when structurally safe. This does not prove numbering/TOC fidelity; numeric section references follow actual resulting numbering capability and are never invented.

### Unordered list item

Strict:

```text
explicit unorderedListItem prototype
-> configured/default unordered list style with valid bullet numbering
-> error
```

YOLO:

```text
usable template bullet-list example/style
-> builtin:unordered-list-v1
-> plain paragraph preserving item content
```

The built-in uses real Word numbering, not literal bullet text.

### Ordered list item

Strict:

```text
explicit orderedListItem prototype
-> configured/default ordered list style with valid ordered numbering
-> error
```

YOLO:

```text
usable template ordered-list example/style
-> builtin:ordered-list-v1
-> plain paragraph preserving item content
```

The built-in uses real single-level decimal numbering, not literal number text.

### Generated table

Strict generated-table style behavior remains.

Prepared-table behavior remains governed by the prepared-table contract.

YOLO may choose another concrete table style if the selected style is unavailable. If none exists, an otherwise valid generated table may be emitted without a table-style reference.

If a prepared-table prototype is unusable but the semantic table is available and placement is unambiguous, YOLO may render a generated table at that placement after emitting a degradation.

### Captions

Numbered captions prefer explicit caption prototypes because the prototype owns real `SEQ` semantics, label text, punctuation and presentation.

Existing plain-caption fallback remains where already allowed.

YOLO may degrade to a plain caption paragraph.

If numeric target semantics cannot truthfully be established, YADG does not invent a number; the reference is handled as visibly unresolved by `AUTHORING-RESILIENCE.md`.

### Inline code

Strict `markdown.codeInline: style` requires the configured `codeInline` character style.

YOLO:

```text
compatible character style with strong code/source/monospace hint
-> ordinary run/plain text
```

Do not choose arbitrary unrelated character styles.

### Value replacement

Existing base-format rule remains: the run containing the first logical tag character is the formatting example.

YOLO never invents a missing value.

## Built-in fallback examples

Required identities:

```text
builtin:plain-v1
builtin:unordered-list-v1
builtin:ordered-list-v1
```

Storage is implementation-owned (code, packaged resource, OOXML fragment, etc.).

Built-ins:

- are never used in strict mode;
- are visually conservative;
- are never reported as template-owned;
- are visible through `inspect template` when they would be selected;
- are covered by both Word and LibreOffice compatibility tests.

YOLO list built-ins may create/import the minimal numbering definitions needed for their own example. This narrow exception does not change strict template-owned numbering behavior.

## `inspect template`

Add:

```text
yadg inspect template [--workspace <path>] [--template <filename.docx>]
```

This presents YADG's semantic view of the template.

At minimum report:

### Template/control structure

- front-matter/control region presence;
- placement tags and type/ID;
- prototype IDs and recognized role;
- prepared-table markers where discoverable;
- other template-owned authoring structures useful for diagnosis.

### Presentation roles

For each relevant role:

- configured selector/prototype;
- strict resolved source/status;
- visible name/internal ID/type for style-backed resolution;
- numbering/outline capability where relevant;
- YOLO candidate/fallback preview if strict fails.

Example:

```text
unordered-list
  configured: ListBullet
  strict: unresolved
  yolo:
    template candidate: "List Bullet 2" (id=ListBullet2, bullet numbering)
    builtin fallback: builtin:unordered-list-v1
```

### Human locations

Entries use story/container + heading context + searchable nearby text + supplemental structural ordinal.

Exact output is human-readable, not a stable machine serialization.

## Human DOCX locator

Prefer recognizable/searchable content over bare ordinals.

Example:

```text
YadgTemplates/report.docx
body > "Risk Model" > "Assumptions" > paragraph 16
near: "Missing values are treated as..."
```

Other story:

```text
YadgTemplates/report.docx
default footer > paragraph 2
near: "Document version: {{value:document-version}}"
```

Prototype:

```text
YadgTemplates/report.docx
prototype "bullet-item"
near: "{{item}}"
```

Excerpt rules:

- collapse whitespace;
- omit control characters;
- keep bounded length;
- avoid dumping long document contents;
- retain enough literal text for Word Find where practical.

Package part, paragraph ordinal and `w14:paraId` may be retained as secondary machine data.

## Relationship to `inspect styles`

`inspect styles` answers:

> Which concrete serialized styles exist and what are their IDs/names/aliases/capabilities?

`inspect template` answers:

> What YADG structures/presentation roles exist, what will strict mode use, and what would YOLO do?

Neither command executes external content producers.

## Compatibility

Existing valid template bindings remain valid.

Strict mode does not gain fuzzy/heuristic matching merely because YOLO exists.

## Non-goals

This spec does not create a GUI/template editor, promise every Word gallery item is addressable, or make built-in YOLO styling a production/template-design substitute.
