# Template Presentation Specification

## Status

Authoritative for YADG 1.1 template-presentation resolution after M0012 corrections.

This supersedes narrower statements that make style lookup or per-command ad hoc resolution the presentation model.

## Principle

YADG is template-first.

Presentation can come from:

- placement/context paragraphs;
- explicit YADG prototypes;
- Word styles;
- numbering definitions;
- prepared table structures;
- placeholder-run formatting;
- template-owned fields/structures;
- explicit built-in YOLO fallback examples.

These are mechanisms under one semantic presentation resolver.

## Single source of truth

Strict validation, actual authoring, `inspect template`, and YOLO preview consume the same semantic presentation result.

They must not independently answer whether a role is resolved.

Implementation structure is free, but behavior must be equivalent to:

```text
ResolvePresentation(role, template, context, mode)
    -> PresentationResolution
```

with enough data to express:

```text
role
strict status
selected source kind
selected source identifier/name
capability
YOLO related candidate
YOLO built-in fallback
diagnostics
```

## Source kinds

A role may resolve through:

```text
prototype
contextual-example
configured-resource
compatibility-default
related-template-resource   # YOLO only
builtin-fallback            # YOLO only
plain-representation        # YOLO only
visible-placeholder         # YOLO only where defined
```

## Strict resolution

Role-specific precedence is authoritative.

Generic shape:

```text
explicit prototype/example
-> contextual example where applicable
-> configured resource
-> role-specific compatibility default
-> error
```

A source earlier in the chain suppresses irrelevant later requirements.

A valid `unorderedListItem` prototype means the unordered-list role is strictly resolved even if the configured/default list style is absent. The same rule applies to ordered-list and caption prototypes and prepared-table/contextual examples.

## YOLO resolution

YOLO first evaluates the complete strict chain.

If strict resolution fails for a recoverable presentation reason:

```text
compatible related template resource
-> versioned built-in fallback
-> plain truthful representation
-> visible placeholder where defined
-> degradation
```

YOLO does not change strict selector matching rules.

## List capability

### Unordered

A usable unordered-list source has effective real Word numbering whose effective level is bullet format.

### Ordered

A usable ordered-list source has effective real Word ordered numbering and is not restricted to decimal.

Roman, alphabetic, decimal variants, and other Word ordered formats qualify when the effective level represents a visible ordered sequence.

Bullet and explicit no-numbering/plain formats do not qualify.

The same capability classifier is used by:

- strict validation;
- related-candidate discovery;
- authoring;
- `inspect template`.

## Role contracts

### Ordinary/contextual paragraph

Preferred source:

```text
contextual placement/insertion paragraph
```

YOLO fallback:

```text
builtin:plain-v1
```

### Heading

Strict:

```text
configured heading resource
-> existing HeadingN compatibility default
-> error
```

A heading resource's effective outline level is resolved using direct and inherited style semantics.

YOLO may select a compatible paragraph style with the required effective outline level.

If none exists and structurally safe, use plain paragraph presentation with direct outline level as a degradation.

### Unordered list

Strict:

```text
unorderedListItem prototype
-> configured/default bullet-capable paragraph style/resource
-> error
```

YOLO:

```text
compatible related bullet resource
-> builtin:unordered-list-v1
-> plain item paragraphs
```

### Ordered list

Strict:

```text
orderedListItem prototype
-> configured/default ordered-numbering paragraph style/resource
-> error
```

YOLO:

```text
compatible related ordered-numbering resource
-> builtin:ordered-list-v1
-> plain item paragraphs
```

### Generated table

Strict uses configured/default table presentation.

YOLO may select another compatible table style, then a structurally valid unstyled table.

### Prepared table

The prepared-table prototype row owns its row presentation.

If the prepared prototype is unusable and the semantic table is unambiguous, YOLO may degrade to generated-table rendering.

Inspection reports whichever source authoring actually uses.

### Caption

Strict:

```text
figure/table caption prototype
-> existing role-specific plain/style fallback where already allowed
-> error when required semantics cannot be established
```

A valid caption prototype owns SEQ semantics and presentation. Missing caption style is irrelevant when the prototype fully owns the role.

YOLO may degrade to a plain caption; numeric semantics are never invented.

### Inline code

Strict:

```text
configured codeInline character style
-> error when policy=style
```

YOLO:

```text
compatible related code/source/monospace character style
-> ordinary text
```

### Value replacement

Existing placeholder-run formatting remains the contextual example.

Missing-value recovery is governed by resilience, not by inventing content.

## Built-in fallbacks

Required identities:

```text
builtin:plain-v1
builtin:unordered-list-v1
builtin:ordered-list-v1
```

List built-ins use real numbering definitions.

They are:

- YOLO-only;
- versioned;
- inspectable;
- clearly identified as implementation-owned, not template-owned;
- tested through Word and LibreOffice.

## `inspect template`

`inspect template` consumes the same presentation resolution used by strict analysis/authoring.

It reports semantic roles, not competing raw mechanisms.

Example with prototype:

```text
unordered-list
  strict: resolved
  source: prototype
  id: bullet-item
  near: "{{item}}"
```

It must not also claim the role is unresolved because `ListBullet` is absent.

Example without prototype:

```text
unordered-list
  configured: ListBullet
  strict: unresolved
  yolo:
    related: "Corporate Bullet" (id=CorpBullet)
    builtin: builtin:unordered-list-v1
```

Raw style inventory remains the job of `yadg inspect styles`.

## Human location context

Presentation inspection and Word diagnostics use the shared structured location model.

For body content, nearest heading path derives from effective outline semantics, including:

- direct paragraph outline level;
- paragraph style outline level;
- inherited `basedOn` outline level;
- equivalent heading semantics already used by document composition.

Literal `Heading1`...`Heading9` IDs are not required.

Visible shape:

```text
YadgTemplates/report.docx
body > "Risk Model" > "Assumptions" > paragraph 16
near: "Missing values are treated as..."
```

CLI keeps the file path workspace-relative.

## Deterministic related-resource selection

Only capability-compatible resources qualify.

Within equally capable candidates:

1. stronger role-specific semantic/name hint;
2. normalized visible name;
3. internal ID.

Do not use incidental document/hash-map enumeration order.

## Compatibility

Existing valid strict templates remain valid.

These corrections remove contradictory duplicate resolution logic; they do not make YOLO heuristics part of strict mode.

## Non-goals

This specification does not create a GUI template editor or make fallback examples a substitute for production template design.
