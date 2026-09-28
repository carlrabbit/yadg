# Authoring Resilience and YOLO Mode Specification

## Status

Authoritative for explicit YOLO recovery introduced by the M0012 correction.

Strict mode remains authoritative when `--yolo` is absent.

## Purpose

YADG must serve two different questions.

Strict:

> Is this workspace/template/runtime production-correct under the configured contract?

YOLO:

> What is the best useful artifact YADG can safely produce from what I currently have?

YOLO is an authoring escape hatch, not silent error suppression.

## Activation

YOLO is invocation-local:

```text
yadg check --yolo
yadg build --yolo
yadg render --yolo
```

M0012 does not persist YOLO in `YADG.md`.

`init`, `inspect styles`, and `inspect template` do not need YOLO.

`publish` remains finalized-artifact-only and has no PreWord fallback.

## Degradation diagnostics

A recovered strict error becomes a visible degradation.

Example:

```text
degradation YADG-YOLO-LIST-001 YadgTemplates/report.docx:
style "ListBullet" is unavailable; fallback: builtin:unordered-list-v1
```

Each degradation identifies:

- original problem;
- selected fallback;
- location/context where available.

Successful YOLO commands summarize the number of degradations.

## Exit semantics

### Check

`check --yolo` succeeds when every otherwise-fatal problem is absent or covered by a defined safe YOLO recovery.

It does not rewrite files.

### Build

`build --yolo` succeeds when no fatal global boundary is crossed and at least one requested template produces a useful authored DOCX.

A recoverable failed template may be skipped while other templates continue; the skip is a degradation.

If no authored DOCX is produced, build fails.

### Render

`render --yolo` succeeds when the requested renderer succeeds or the alternate supported renderer successfully finalizes all authored inputs for that invocation.

If both fail, render fails.

`YadgPreWords` is never promoted to finalized output merely to make YOLO succeed.

### Publish

Publishing remains:

```text
YadgWords/*.docx
```

only.

A DOCX finalized by an alternate renderer under YOLO is still finalized and publishable.

## Recovery catalog

### Presentation resources

Use `TEMPLATE-PRESENTATION.md`.

YOLO may select a compatible related template resource, use a built-in fallback example, or preserve content plainly.

### Thematic break

If strict policy would error, YOLO may omit the thematic break and emit a degradation.

### Inline code

If configured presentation cannot be satisfied:

1. use a compatible related code/source/monospace character style if safely identifiable;
2. otherwise preserve code content as ordinary text.

Never drop the code content.

### Missing workspace value

Strict remains error.

YOLO preserves the original visible token:

```text
{{value:missing-id}}
```

and emits a degradation.

It does not invent empty/guessed content.

### Unresolved semantic reference

Strict remains error.

YOLO preserves an obvious visible unresolved representation equivalent to the source identity, e.g.:

```text
[@missing-id]
```

It does not invent a target or number.

Duplicate/ambiguous target identity is fatal because choosing one would fabricate meaning.

### Missing ordinary figure asset

If the path itself is valid and inside the workspace but the file is absent/unreadable, YOLO may emit a visible placeholder containing:

- stable figure ID;
- caption/alt text where available;
- short failure reason.

Example:

```text
[YADG figure "system-context" unavailable: images/system-context.png]
```

Path escape/reparse/security violations remain fatal.

### External producer failure

If Mermaid/external producer execution fails/unavailable, YOLO may emit a visible placeholder semantic figure rather than dropping the object.

The producer failure remains visible as a degradation.

YOLO does not substitute an arbitrary unconfigured executable.

### List presentation

Resolution follows:

```text
strict prototype/style chain
-> compatible real template list resource
-> builtin unordered/ordered list fallback
-> plain item paragraphs
```

Built-in list fallbacks use real numbering.

Plain fallback preserves item content.

### Caption numbering

YOLO may render a plain caption if numbered prototype mechanics are unavailable.

A numeric reference that can no longer be established remains visibly unresolved rather than receiving an invented number.

### Individual unreadable template

When multiple templates exist, YOLO may skip a template for a file-local failure that can be isolated safely and continue others.

A corrupt package is never modified/guessed through.

If no useful authored output remains, build fails.

## Renderer fallback

### Requested Word

```text
word
-> failure/unavailable under --yolo
-> libreoffice
```

### Requested LibreOffice

```text
libreoffice
-> failure/unavailable under --yolo
-> word
```

If an explicit LibreOffice `--renderer-path` fails and Word is substituted, report both facts.

### Retry safety

Preserve existing renderer staging/commit behavior.

Attempt fallback only after the primary renderer reports failure without committing successful-looking final output for that invocation.

Do not merge partial output from primary and fallback runs.

### Reporting

Preserve primary diagnostics as degradation context.

Successful fallback output must identify:

```text
requested renderer
actual renderer
actual runtime version
YOLO fallback = yes
```

Do not print the requested renderer as if it actually rendered the document.

### Both renderers fail

Return non-zero.

Leave PreWords untouched as authored intermediates.

Do not put them into `YadgWords`.

## Fatal boundaries

YOLO does not guess through:

- malformed workspace configuration that cannot be interpreted;
- duplicate/ambiguous semantic identity;
- unsafe path traversal/reparse/security conditions;
- corrupt DOCX state that prevents safe package handling;
- a recovery that would fabricate semantic content;
- inability to produce any requested useful output;
- both renderers failing;
- unsafe publication destination;
- external publication/authentication failures.

Rule:

> YOLO may sacrifice presentation fidelity or runtime choice. It may not fabricate identity, bypass safety, or misstate what happened.

## Determinism

Fallback choice is deterministic.

Use the candidate eligibility/tie-break rules in `TEMPLATE-PRESENTATION.md`.

Never define "first" by incidental enumeration order.

## Inspection

`yadg inspect template` previews relevant YOLO presentation fallbacks without changing the template.

`yadg check --yolo` reports recoveries the current workspace would require.

## Documentation

Show the intended workflow:

```powershell
# production correctness
yadg check
yadg build

# exploratory/best effort
yadg check --yolo
yadg build --yolo
yadg render --renderer word --yolo
```

Explain that production/CI normally uses strict commands.

YOLO artifacts are useful for authoring progress; degradation diagnostics identify what remains to fix.

## Validation rule

For each recoverable class, pair strict and YOLO evidence:

```text
strict -> fails
YOLO   -> succeeds with explicit degradation
```

Fatal scenarios fail in both modes.

Real cross-renderer fallback is validated on the declared Windows/Word/LibreOffice Tier-3 locus.
