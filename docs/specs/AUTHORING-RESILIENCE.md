# Authoring Resilience and YOLO Mode Specification

## Status

Authoritative for YADG 1.1 resilience after M0012 corrections.

Strict mode remains default and authoritative when `--yolo` is absent.

## Purpose

Strict asks:

> Is this production-correct?

YOLO asks:

> What is the best useful artifact YADG can safely and truthfully produce now?

YOLO is not silent exception swallowing.

## Trusted local workspace

YADG is a local user-scope authoring tool. The workspace selected by the user is trusted input; YADG is not a hostile-filesystem sandbox.

Accordingly:

- ordinary file symlinks/reparse points are not fatal merely for being links;
- the workspace root and explicitly addressed conventional directories may be reached through links/junctions;
- recursive discovery does not follow linked/reparse child directories, to avoid cycles/duplicate/unbounded traversal;
- linked recursive directories may be skipped with a non-fatal warning;
- explicit asset paths that lexically escape the workspace remain invalid for portability/reproducibility;
- YADG does not resolve final symlink targets merely to establish physical containment.

The inability of a Windows test account to create symlinks is not an acceptance failure.

## Activation

Invocation-local only:

```text
yadg check --yolo
yadg build --yolo
yadg render --yolo
```

No persistent YOLO mode is introduced.

## Degradation model

A recovered strict failure becomes one explicit logical degradation event.

Each event contains:

- code;
- original problem;
- selected fallback;
- location/context where available;
- optional detail diagnostics.

Command summaries count logical degradation events consistently.

## Template failure classification

Do not implement YOLO template handling as:

```text
catch any exception -> skip template
```

Failures are classified.

A file-local condition may be skipped under YOLO when it can be isolated safely, another usable template remains, and skipping requires no semantic guess.

Corrupt/unreadable state that prevents safe DOCX handling and unknown unexpected implementation failures remain errors rather than being silently normalized to a skip.

## `check --yolo`

Check succeeds only if:

- all remaining issues are either non-fatal or safely recovered; and
- at least one template is usable for a potential build.

If every template is skipped/unusable, check fails explicitly.

## `build --yolo`

Build succeeds only if at least one requested template produces an authored DOCX and no fatal global boundary is crossed.

A recoverable template may be skipped.

### Complete PreWord result set

On successful strict or YOLO build:

```text
YadgPreWords/*.docx
```

represents exactly the outputs produced by that invocation.

Stale prior outputs do not survive.

Use staging/commit or equivalent result-set replacement. If build fails before commit, do not leave a partially new set.

## Recovery catalog

Retain the defined recoveries:

- presentation-resource fallback through `TEMPLATE-PRESENTATION.md`;
- thematic-break omission;
- inline-code related style/plain text;
- missing value token preserved visibly;
- unresolved reference preserved visibly;
- missing in-workspace figure asset -> visible placeholder;
- Mermaid producer failure -> visible placeholder;
- missing/unusable list presentation -> related resource/built-in/plain;
- caption semantics -> plain caption without invented numbering;
- recoverable single-template skip when another useful output remains.

## Fatal boundaries

Still fatal:

- malformed required configuration that cannot be interpreted for the operation;
- duplicate/ambiguous semantic identity;
- explicit lexical workspace asset escape;
- corrupt DOCX state that prevents safe handling;
- recovery requiring fabricated semantic meaning;
- no usable template/output;
- both supported renderers failing;
- false claims that an operation succeeded.

A symlink/reparse attribute by itself is not a fatal boundary.

## Renderer fallback

Under explicit `--yolo`:

```text
word -> libreoffice
libreoffice -> word
```

Each attempt uses staged output.

The actual renderer identity is explicit result provenance, not inferred from incidental fields.

Successful fallback reports:

```text
requested renderer
actual renderer
actual runtime version
YOLO fallback = yes
```

A renderer substitution is one logical degradation event. Primary renderer diagnostics may be detail lines beneath that event.

If both renderers fail, return non-zero and do not promote PreWords.

## Finalized result-set freshness

A successful render owns the complete generated finalized result set.

For actual LibreOffice:

```text
YadgWords/*.docx
YadgPdfs/*.pdf
```

correspond exactly to current PreWords.

For actual Word:

```text
YadgWords/*.docx
```

correspond exactly to current PreWords and stale PDFs from earlier LibreOffice runs are absent.

This includes YOLO alternate-renderer success.

Use result-set staging/commit so prior finalized artifacts are not deleted merely because a new renderer attempt started.

## Publishing boundary

Publishing is downstream of authoring.

`publish` does not use the authoring pipeline, parse Markdown, inspect templates, validate figures, or execute content producers.

With explicit `--publish-path`, current authoring validity is irrelevant to an already-finalized `YadgWords` artifact.

Without the CLI destination, only the configuration needed to obtain `publish.path` is read.

## Relative/human diagnostics

CLI workspace-local paths are relative.

DOCX context is structured separately from the filesystem path and includes searchable nearby text/effective heading context where available.

Do not make relative-path conversion depend on parsing a multiline location string.

## Discovery pruning

Recursive Markdown discovery does not descend into:

```text
YadgTemplates
YadgPreWords
YadgWords
YadgPdfs
dot-prefixed directories
linked/reparse child directories
```

The first five categories are deliberately excluded source trees. Linked child directories are skipped to avoid recursive filesystem graph traversal.

## Validation rules

For each recoverable class, pair strict and YOLO evidence where applicable.

For each true fatal boundary, prove YOLO still fails.

Tests also prove:

- ordinary link attributes are not rejected as fatal;
- link creation privilege is not a product prerequisite;
- recursive linked directories are skipped non-fatally when testable;
- stale PreWords and finalized outputs do not survive successful result-set replacement;
- publisher independence from authoring/producers;
- actual renderer provenance/degradation counting.

## Tier-3 fixture rule

Recovery scenarios whose acceptance relies on Word-origin template behavior use committed templates actually created/saved through Microsoft Word.

Direct XML/ZIP mutation may be useful for focused synthetic tests but is not authoritative Word-origin Tier-3 evidence.

## Documentation

Keep the intended workflow clear:

```text
strict = production correctness
YOLO   = truthful best-effort authoring progress
```

YOLO may reduce fidelity. It may not misstate output freshness, selected renderer, semantic identity, or success.
