# Authoring Experience Specification

## Status

Authoritative for YADG 1.1 authoring ergonomics introduced by M0012.

This specification covers workspace bootstrap, supported configurable Markdown policies, CLI discovery/check behavior, and diagnostic presentation.

Template style discovery/binding and list prototypes are specified separately in `docs/specs/TEMPLATE-STYLES.md`.

## Design goal

A YADG author should not need knowledge of Markdig implementation class names, absolute filesystem paths, or Word's internal style identifiers to perform ordinary authoring work.

The CLI should help users establish a workspace, understand what YADG discovered, locate errors in their source, and discover the template presentation vocabulary available to them.

## Workspace bootstrap

YADG 1.1 adds:

```text
yadg init [--workspace <path>]
```

`--workspace` follows the same default as other workspace commands: current directory when omitted.

The target directory may already contain unrelated files. `init` does not require an empty directory.

Before mutation, `init` checks all paths it owns. It must not overwrite or merge an existing YADG workspace or starter file.

The owned bootstrap paths are:

```text
YADG.md
content.md
YadgTemplates/
```

If any owned bootstrap path already exists, `init` fails before creating/changing any of them and reports the conflicts.

If the target directory does not exist, `init` may create it.

Successful `init` creates:

- `YADG.md` with valid schema-v1 front matter;
- `content.md` with one starter stable section;
- empty `YadgTemplates/`.

It does not synthesize a DOCX template and does not create generated-output directories.

The starter Markdown section ID is:

```text
introduction
```

The success message tells the user to add a prepared DOCX template containing, at minimum, a placement such as:

```text
{{content:introduction}}
```

and then run `yadg check`.

Generated workspace directories such as `YadgPreWords`, `YadgWords`, and `YadgPdfs` are created by the commands that own them and their absence is not a `check` warning.

M0012 does not add `--force`.

## Markdown policy configuration

`YADG.md` schema version remains 1.

The allowed root mapping is extended with optional:

```yaml
markdown:
  thematicBreak: error
  codeInline: error
```

The supported values are:

```text
thematicBreak: error | ignore
codeInline:     error | ignore | style
```

Unknown keys or values are configuration errors.

The defaults are `error`, preserving 1.0 behavior for existing workspaces.

These keys are stable YADG product vocabulary. Markdig implementation type names such as `ThematicBreakBlock` and `CodeInline` are diagnostic/internal concepts and are not configuration keys.

### Thematic break

With:

```yaml
markdown:
  thematicBreak: ignore
```

a Markdown thematic break is deliberately omitted from semantic/rendered content and does not produce an unsupported-construct diagnostic.

`error` retains the unsupported-construct failure.

M0012 does not define a rendered horizontal-rule Word representation.

### Inline code

With:

```yaml
markdown:
  codeInline: ignore
```

the code span's textual content is preserved as ordinary text while the Markdown code styling semantics are ignored.

The content is never silently dropped.

With:

```yaml
markdown:
  codeInline: style
```

inline code becomes a semantic inline that is rendered using the template-owned `codeInline` character-style role defined by `docs/specs/TEMPLATE-STYLES.md`.

A template only needs that role when rendered/selected content actually contains styled inline code.

`error` retains the 1.0 unsupported-inline failure.

Inline code remains inline content; fenced non-Mermaid code blocks remain unsupported by M0012.

## Diagnostics

### Workspace-relative paths

Diagnostics for files inside the active workspace display workspace-relative paths.

Example:

```text
content.md
YadgTemplates/report.docx
docs/model.md
```

Normal diagnostics do not print the workspace's absolute path prefix.

Paths outside the workspace, when a diagnostic legitimately refers to an external resource, are not falsely relativized.

### Text source positions

Diagnostics originating from a concrete location in Markdown or `YADG.md` include 1-based line and column when the parser can identify the responsible token/node.

Canonical human display shape:

```text
<severity> <code> <relative-path>:<line>:<column>: <message>
```

Example:

```text
error YADG-MD-UNSUPPORTED docs/model.md:37:14: Unsupported Markdown inline 'CodeInline'.
```

Global/cross-file diagnostics that do not correspond to one textual token may omit line/column but retain the relative path where applicable.

### DOCX structural positions

DOCX has no user-meaningful source line number.

Template diagnostics use the relative DOCX path plus a structural locator sufficient to find the relevant Word location when reasonably available, for example:

```text
YadgTemplates/report.docx [body paragraph 16]
YadgTemplates/report.docx [default footer paragraph 2]
YadgTemplates/report.docx [prototype bullet-item]
```

M0012 does not invent fake line numbers for DOCX.

### Severity and output

`check` prints errors and warnings.

Exit status remains failure only when errors are present.

Diagnostic codes remain stable troubleshooting identifiers; message prose and structural detail may improve.

## Check summary and discovery listing

The existing success summary remains, with reference count defined as all referenceable semantic objects:

```text
sections + tables + figures
```

Example:

```text
check: valid (1 source(s), 1 template(s), 4 reference(s))
```

M0012 adds:

```text
yadg check --list [--workspace <path>]
```

`--list` prints deterministic, workspace-relative discovery sections for:

- Markdown sources;
- DOCX templates;
- referenceable semantic objects.

Referenceable objects are listed with stable ID and type:

```text
section
table
figure
```

Example shape:

```text
Sources
  docs/content.md

Templates
  YadgTemplates/report.docx

References
  architecture        section
  interface-matrix    table
  system-flow         figure

check: valid (1 source(s), 1 template(s), 3 reference(s))
```

Ordering is deterministic.

When the workspace can be discovered but is invalid, `--list` still lists the successfully discovered items and diagnostics, then reports an invalid summary/exit status.

`--list` is a human-readable authoring aid, not a machine-stable serialization format.

## Style discovery command

The CLI adds:

```text
yadg inspect styles [--workspace <path>] [--template <filename.docx>]
```

`inspect styles` is defined in detail by `docs/specs/TEMPLATE-STYLES.md`.

The command is intentionally usable when ordinary `check` fails because of style/list configuration. It must not require the complete Markdown workspace to validate successfully before it can inspect a DOCX template.

## Documentation contract

YADG 1.1 adds dedicated user documentation rather than expanding README into the complete manual.

At minimum:

```text
docs/user/GETTING-STARTED.md
docs/user/AUTHORING.md
docs/user/TEMPLATES.md
docs/user/TROUBLESHOOTING.md
```

README remains the concise product/install/workflow landing page and links these documents.

The user documentation must cover:

- `yadg init`;
- workspace/source discovery;
- supported Markdown and unsupported/configurable constructs;
- stable IDs, sections/content, references;
- figures/tables/prepared tables;
- values and Mermaid;
- template control/front matter;
- style discovery and style-selector rules;
- list style/prototype approaches;
- inline-code style configuration;
- real examples of common `check` failures and fixes;
- renderer/publish workflow;
- diagnostics and `check --list`.

The templates documentation must explain the distinction between Word UI style names, aliases, internal style IDs, hidden/semi-hidden styles, and numbering/list definitions without requiring the user to edit raw OOXML.

## Compatibility

Existing 1.0 workspaces and templates without the new optional configuration retain their previous semantics.

`check` output becomes more informative; no machine-stable text-output compatibility is promised by 1.0.

M0012 does not introduce a JSON diagnostic/listing format.

## Non-goals

M0012 does not:

- render thematic breaks;
- add fenced code-block rendering other than existing Mermaid;
- add arbitrary per-Markdig-node ignore rules;
- add nested lists;
- add machine-readable `check` output;
- add an interactive TUI/template editor;
- synthesize a starter DOCX template;
- change renderer or publishing semantics.

## Explicit YOLO recovery (M0012 correction)

`check`, `build`, and `render` accept invocation-local `--yolo`; strict behavior remains the default. A recovered error is printed as `degradation YADG-YOLO-*`, including the original problem and selected fallback. Successful YOLO commands report a degradation count. YOLO can substitute deterministic compatible presentation, versioned built-ins, visible placeholders, preserved unresolved tokens, or the other supported renderer. It never changes `YADG.md` or weakens malformed-config, duplicate-ID, path/reparse, corrupt-package, no-output, both-renderers-failed, or unsafe-publication boundaries. `publish` remains finalized-DOCX-only.

`yadg inspect template [--workspace <path>] [--template <filename.docx>]` explains control/front-matter structures, placements and searchable locations, prototypes, strict role resolution, and YOLO fallback candidates. It does not run producers. See `TEMPLATE-PRESENTATION.md` and `AUTHORING-RESILIENCE.md`; they supersede narrower role/failure rules below where explicitly stated.
