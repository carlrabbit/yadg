# Workspace Values Specification

## Status

Authoritative for root `YADG.md` workspace front matter and scalar values. M0008 extends the same front-matter schema with external producer configuration defined separately in `docs/specs/CONTENT-PRODUCERS.md`.

## `YADG.md`

`YADG.md` remains the workspace marker. It may contain optional YAML front matter at the start followed by human-facing workspace notes that are not document source.

A file without front matter remains valid and defines empty values and no producers.

## Front-matter syntax

Front matter starts at the first effective line with exactly `---` and ends at the next line exactly `---`. Optional UTF-8 BOM is permitted before the opening delimiter.

An unclosed front matter region is an error.

## Schema version

The front-matter schema remains:

```yaml
yadg:
  version: 1
```

M0008 does not increment that schema version.

Allowed root keys are now:

```text
yadg
values
producers
```

`yadg` is required whenever front matter is present.

`values` is governed by this specification.

`producers` is governed by `docs/specs/CONTENT-PRODUCERS.md`.

Unknown root keys, duplicate mapping keys, aliases, anchors, merge keys, and custom YAML tags remain errors.

## Workspace values

`values` optionally maps value IDs to YAML string scalars.

Value IDs:

```text
[A-Za-z][A-Za-z0-9_-]*
```

They are case-sensitive and occupy a namespace separate from semantic object IDs.

Values are single-line strings only. Numeric, boolean, null, sequence, mapping, and multiline values are invalid. Empty string is valid.

`{{value:<id>}}` resolves values only. `[@id]` resolves semantic objects only.

Missing referenced values are errors. Unused values are allowed.

## Literal inline substitution

Value replacement remains literal/non-recursive.

Supported Word locations remain ordinary paragraphs in:

- main document body;
- main-body table cells;
- headers;
- footers.

Replacement inherits run properties from the run containing the first logical tag character.

Unsupported text-box/note/comment/property/field-instruction locations remain out of scope.

## Security

Values remain source-controlled data, not secret-store/environment/CLI inputs.

Producer configuration under `producers` is executable tooling configuration and has a separate trust boundary defined in `docs/specs/CONTENT-PRODUCERS.md`.

## Compatibility

Existing M0006 front matter containing only `yadg` and `values` remains valid.

M0008 merely adds optional `producers`; it does not change value semantics.
