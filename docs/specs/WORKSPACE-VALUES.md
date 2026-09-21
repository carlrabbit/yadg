# Workspace Front Matter and Values Specification

## Status

Authoritative for root `YADG.md` front matter and workspace scalar values.

Other root mappings are specialized by:

- `producers` -> `docs/specs/CONTENT-PRODUCERS.md`
- `publish` -> `docs/specs/PUBLISHING.md`

Document-wide visible-text replacement behavior is specialized by:

- `docs/specs/DOCUMENT-COMPOSITION.md`

## `YADG.md`

`YADG.md` remains the workspace marker. It may contain optional YAML front matter followed by human-facing notes that are not document source.

A file without front matter remains valid and defines empty values, no producers, and no publication default.

## Front-matter syntax

Front matter starts at the first effective line with exactly `---` and ends at the next line exactly `---`.

An optional UTF-8 BOM is permitted before the opening delimiter.

An unclosed region is an error.

## Schema version

Schema remains:

```yaml
yadg:
  version: 1
```

M0010 does not increment it.

Allowed root keys remain:

```text
yadg
values
producers
publish
```

`yadg` is required whenever front matter is present.

Unknown root keys, duplicate mapping keys, aliases, anchors, merge keys, and custom YAML tags remain errors.

## Values

`values` optionally maps case-sensitive IDs to single-line YAML string scalars.

Value IDs:

```text
[A-Za-z][A-Za-z0-9_-]*
```

Values remain in a namespace separate from semantic object IDs.

Numeric, boolean, null, sequence, mapping, and multiline values are invalid. Empty string is valid.

`{{value:<id>}}` resolves workspace values; `[@id]` resolves semantic objects.

Missing referenced values fail; unused values are allowed.

## Literal substitution

Value replacement is literal and non-recursive.

Replacement inherits formatting from the run/text position containing the first logical tag character.

After M0010, value substitution is a visible-template-text mechanism across the supported Word stories defined by `docs/specs/DOCUMENT-COMPOSITION.md`, including:

- main body and table cells;
- headers;
- footers;
- footnotes;
- endnotes;
- comments;
- supported text-box/shape text.

A value tag may be split across runs within one supported paragraph/text-container scope.

Field instruction code, package metadata/properties, relationship targets, arbitrary XML attributes, and other non-visible metadata are not value-substitution targets.

## Publication mapping

`publish` is optional and governed by `docs/specs/PUBLISHING.md`.

Its `path` is a publication default, not a redefinition of workspace output directories.

`check`, `build`, and `render` remain valid without it.

## Security

Values are source-controlled data, not secrets.

Producer configuration remains trusted executable configuration.

Publication configuration is filesystem destination configuration and defines no credential or remote protocol.

## Compatibility

Existing schema-v1 workspaces containing only `yadg`, `values`, and/or `producers` remain valid.

M0010 changes where existing `{{value:<id>}}` tags are supported; it does not add new workspace value syntax.
