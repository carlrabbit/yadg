# Troubleshooting

## Start with `check`

Run:

```powershell
yadg check --list
```

This validates the workspace and shows what YADG discovered.

Strict errors fail the command. Warnings do not.

Use YOLO only when you deliberately want a best-effort authoring artifact while fixing the strict problem.

## Find a Markdown/YAML error

Source diagnostics use workspace-relative paths and one-based line/column when a concrete location is available.

Example shape:

```text
content.md:18:7
```

## Find a Word-side error

DOCX does not have useful source line numbers.

YADG reports human-oriented context where available:

```text
YadgTemplates/report.docx
body > "Risk Model" > paragraph 16
near: "Missing values are treated as..."
```

Use the `near:` text in Word Find.

For deeper inspection:

```powershell
yadg inspect template
```

## A style name does not resolve

Run:

```powershell
yadg inspect styles
```

Word distinguishes:

- internal style ID;
- primary display name;
- aliases;
- style type;
- visibility;
- numbering relationships.

A style shown in the Word gallery may not exist as a concrete serialized style until it has actually been used/materialized.

YADG strict selector matching is exact:

1. internal ID;
2. primary name;
3. alias.

It is not fuzzy.

## `List style '...' does not resolve to a valid numbering definition`

A paragraph style name does not necessarily imply real list numbering.

Use:

```powershell
yadg inspect styles
yadg inspect template
```

Then either:

- choose a concrete bullet/ordered style whose effective numbering is valid; or
- use a real list-item prototype in the template.

Ordered numbering is not restricted to decimal; valid Roman/alphabetic numbering is allowed.

## Use a list prototype instead

Create a real Word bulleted/numbered paragraph containing:

```text
{{item}}
```

inside:

```text
{{yadg:prototype:<id>}}
...
{{/yadg:prototype:<id>}}
```

and bind it as `unorderedListItem` or `orderedListItem`.

This is normally the most reliable way to preserve template-specific list presentation.

## Unsupported thematic break

Strict default:

```text
error
```

To ignore it:

```yaml
markdown:
  thematicBreak: ignore
```

YOLO may also omit the thematic break for that invocation.

## Inline code is unsupported

Choose one policy:

```yaml
markdown:
  codeInline: error
```

```yaml
markdown:
  codeInline: ignore
```

or:

```yaml
markdown:
  codeInline: style
```

`style` requires a configured template character style:

```yaml
styles:
  codeInline: "Code"
```

YOLO can degrade missing inline-code presentation to a related compatible character style or ordinary text.

## Missing value

Strict mode reports an error for an unresolved:

```text
{{value:id}}
```

Check `YADG.md` spelling/value configuration.

YOLO preserves the unresolved token visibly; it does not invent an empty or guessed value.

## Missing semantic reference

For:

```text
[@id]
```

check:

- ID spelling;
- duplicate IDs;
- whether the target is rendered in this template;
- whether required numbering/caption semantics actually exist.

YOLO may preserve an unresolved reference visibly, but duplicate/ambiguous semantic identity remains fatal.

## Missing figure

Check:

- path is relative to the Markdown source;
- file exists/readable;
- PNG/JPEG format;
- path remains lexically inside the workspace.

Remote/data URI assets are unsupported.

YOLO can insert an obvious placeholder for a missing/unreadable in-workspace figure; it does not silently remove the semantic object.

## Mermaid failure

Producer configuration is executable-code configuration.

Check:

- executable/path;
- configured arguments;
- package-manager availability;
- producer stderr;
- Mermaid syntax;
- output PNG generation.

YADG does not install Mermaid tooling automatically.

Under YOLO, a failed producer can become an obvious placeholder figure.

## Word render failure

Requirements:

- Windows;
- installed/activated desktop Word;
- interactive logged-on user session;
- `Word.Application` registered;
- no blocking modal prompt.

YADG uses late-bound COM and does not require Office interop DLLs.

On a non-Windows runtime, a Word-render request should fail with a normal actionable YADG platform diagnostic rather than an uncaught platform exception.

## LibreOffice render failure

Install Writer and, when needed, provide:

```powershell
yadg render --renderer libreoffice --renderer-path "C:\Program Files\LibreOffice\program\soffice.com"
```

If the requested renderer fails under YOLO, YADG may try the other supported renderer.

A successful fallback reports requested and actual renderer.

## Renderer YOLO example

```powershell
yadg render --renderer word --yolo
```

Conceptually:

```text
requested renderer: word
Word unavailable

degradation: renderer fallback
actual renderer: libreoffice
```

The exact diagnostic wording is human-readable rather than a machine-stable API.

If both renderers fail, rendering fails. YADG never treats `YadgPreWords` as finalized.

## `check --yolo` still fails

That is expected when no truthful useful recovery exists.

Examples that remain fatal include:

- malformed required configuration;
- duplicate/ambiguous semantic IDs;
- lexically escaping figure paths;
- corrupt DOCX state that prevents safe handling;
- no usable template/output.

YOLO means best useful artifact, not unconditional success.

## File links / symlinks

YADG treats the explicitly selected local workspace as trusted.

An ordinary file is not rejected just because it is a symlink/reparse point.

Recursive Markdown discovery does not follow linked child directories, which avoids loops/duplicate traversal.

A Windows account not permitted to create symlinks is not missing a YADG runtime prerequisite.

## Generated output looks stale

Successful `build` replaces the complete current:

```text
YadgPreWords/*.docx
```

Successful `render` replaces current finalized generated outputs:

```text
YadgWords/*.docx
YadgPdfs/*.pdf
```

Word success clears stale PDFs from earlier LibreOffice output.

If a command fails before committing its new result set, YADG should not present partial new generated output as successful.

Keep unrelated persistent files outside these generated directories.

## Publish fails despite an already-finalized document

With an explicit destination:

```powershell
yadg publish --publish-path ./Published
```

publication is downstream-only. It should not parse current Markdown/templates or execute Mermaid.

Check:

- `YadgWords/` exists;
- at least one top-level finalized `.docx` exists;
- destination is writable;
- destination is not a reserved YADG input/generated directory.

Without `--publish-path`, verify:

```yaml
publish:
  path: ./Published
```

in `YADG.md`.

## When to use strict vs YOLO

Use strict mode to decide whether the document is production-correct:

```powershell
yadg check
yadg build
yadg render --renderer word
```

Use YOLO during exploration when a visible degraded artifact is more useful than being blocked:

```powershell
yadg check --yolo
yadg build --yolo
yadg render --renderer word --yolo
```

Fix the reported degradations before treating the workspace as strict-production-ready.
