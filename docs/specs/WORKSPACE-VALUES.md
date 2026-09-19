# Workspace Values Specification

## Status

Authoritative for workspace-level scalar values and `{{value:<id>}}` substitution introduced by M0006.

## Purpose

Workspace values represent document/workspace facts independent of Word presentation, such as document version, reporting date, owner, or classification.

Values are sourced from root `YADG.md`. Word templates remain authoritative for visual presentation.

## `YADG.md`

`YADG.md` remains the workspace marker. M0006 adds optional YAML front matter at the start of the file.

The Markdown body after front matter remains human-facing workspace notes and is not document source content.

A `YADG.md` without front matter remains valid and defines an empty value set.

## Front-matter syntax

When present, front matter starts at the first effective line with exactly `---` and ends at the next line containing exactly `---`. An optional UTF-8 BOM does not prevent detection.

If the first effective line is not `---`, the file has no machine-readable front matter. An opening delimiter without a closing delimiter is an error.

Example:

```yaml
---
yadg:
  version: 1

values:
  document-version: "2.3"
  reporting-date: "2026-09-30"
  owner: "Liquidity Risk"
---

# Workspace notes
This body is for humans and is not authored into the document.
```

## Schema

M0006 defines schema version `1`.

The root keys are exactly `yadg` and optional `values`.

`yadg` is required when front matter is present and contains exactly:

```yaml
version: 1
```

`values` maps value IDs to scalar string values.

Unknown keys, duplicate mapping keys, aliases, anchors, merge keys, and custom YAML tags are errors.

## Value IDs and namespace

Value IDs use:

```text
[A-Za-z][A-Za-z0-9_-]*
```

They are case-sensitive.

Values occupy a namespace separate from semantic object IDs. A semantic object and a value may therefore use the same spelling.

`[@id]` resolves only semantic objects. `{{value:id}}` resolves only workspace values.

## Value type

M0006 values are YAML string scalars only and must not contain CR or LF characters.

Numeric, boolean, null, sequence, and mapping values are invalid rather than implicitly converted. The empty string is valid.

Date-like or numeric-looking values should be quoted where necessary to ensure YAML treats them as strings.

## Literal substitution

Template syntax is:

```text
{{value:<value-id>}}
```

Replacement is the exact configured string.

Value contents are literal and are not recursively interpreted as YADG tags, Markdown, Word fields, environment-variable expressions, or template syntax.

A missing referenced value is a `check` error. Unused defined values are allowed.

## Inline behavior

`value` is an inline template tag. `content`, `section`, `table`, and `figure` remain block tags.

A paragraph may contain ordinary text and one or more value tags:

```text
Document version {{value:document-version}} — as of {{value:reporting-date}}
```

The tag may be split across OOXML runs.

Replacement preserves unrelated paragraph/run content.

The replacement inherits the run properties of the run containing the first character of the logical tag. If the tag spans differently formatted runs, the replacement is inserted at the first-tag-character position using that first run's formatting; tag characters are removed from later participating runs without merging their formatting.

## Supported Word locations

M0006 substitution is supported in ordinary WordprocessingML paragraphs in:

- the main document body;
- table cells in the main document body;
- headers;
- footers.

Multiple occurrences are allowed.

Block authoring tags retain their existing location restrictions.

M0006 does not support value substitution in text boxes/drawing text, footnotes, endnotes, comments, document/custom properties, field instructions, or template control/prototype regions.

Observable value tags in unsupported locations must be diagnosed rather than silently treated as supported.

## Interaction with template metadata

Template front matter/prototypes remain template-local presentation metadata.

Workspace values do not override template front matter, and template front matter does not define workspace values.

## Interaction with Markdown

M0006 does not interpolate workspace values into Markdown source.

A literal `{{value:id}}` in Markdown remains ordinary Markdown text under the existing Markdown rules.

## Validation

`yadg check` validates `YADG.md` front matter before normal output modification.

Validation covers delimiter correctness, schema version, allowed/duplicate keys, unsupported YAML features, value-ID syntax, string/single-line type, template value resolution, and observable unsupported locations.

`build` performs equivalent validation before changing normal authored outputs.

## Security and reproducibility

M0006 values are source-controlled workspace data.

M0006 does not read values from environment variables, CLI overrides, secret stores, external files, network services, or template-local fallbacks.

Secrets must not be placed in `YADG.md` merely because values are supported.

## Compatibility

Existing note-only `YADG.md` workspaces remain valid. Existing templates without value tags are unaffected.

The previously reserved `value` vocabulary becomes implemented according to this specification.

## Deferred behavior

M0006 does not define typed values, locale formatting, multiline values, override layers, secret references, computed values, Markdown interpolation, text-box/note/comment/property substitution, per-template value namespaces, or conversion of values to Word fields.
