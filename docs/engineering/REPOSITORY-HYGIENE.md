# Repository Hygiene

## Status

Authoritative repository-content policy introduced by M0012.

## Product truth vs operational state

The source repository tracks durable product/project truth such as:

- source and tests;
- engineering scripts;
- product specifications/architecture/terminology;
- milestone contracts/history;
- public/user documentation;
- committed test fixtures whose provenance is part of validation;
- changelog/license/package metadata.

Implementation/review/generated operational state is not product truth and is not tracked on the normal product branch.

The following paths are local/generated operational state:

```text
.execution/
.review/
artifacts/
```

They are ignored by Git after M0012.

## Execution ledgers

The guide-driven implementation process still uses:

```text
.execution/<milestone>.md
```

as persistent local execution state during implementation.

It remains implementation-owned and required when the active milestone/execution process requires it, but it is not committed as product content.

Historical ledgers already present in Git history remain available through history; M0012 removes them from the current tracked tree.

## Review state

Future milestone review requests/records may exist locally under `.review/` when a milestone requires human review.

They are milestone-operational evidence rather than normal product documentation and are not committed after M0012 unless a future milestone explicitly establishes a durable project artifact that belongs elsewhere.

Review tooling must continue to function with ignored local `.review/` state.

## Generated artifacts

`artifacts/` is ignored and removed from the tracked tree.

Generated packages, release evidence, review outputs, temporary documents, and similar outputs are not committed merely because validation produced them.

Durable release information belongs in normal release surfaces such as version authority, changelog, tags/releases/feed metadata, or a deliberately specified durable text record.

## Test fixtures are not generated artifacts

Committed fixtures under test/source-controlled fixture locations remain tracked when they are required to reproduce product validation.

The `artifacts/` rule must not accidentally remove realistic Word/LibreOffice fixtures introduced by earlier milestones.

## Specification organization

`docs/SPECS.md` is the specification index and cross-cutting map.

Detailed normative behavior lives in focused documents under:

```text
docs/specs/
```

`docs/SPECS.md` must link every active focused specification and state important specialization/supersession relationships rather than duplicating partial stale versions of those specs.

## Release-history reconciliation

M0012 starts after YADG 1.0.0 has been released.

Repository truth must therefore no longer describe M0011 as merely `ready` or 1.0.0 as an unreleased candidate.

`CHANGELOG.md` keeps a fresh `[Unreleased]` section for post-1.0/1.1 development and a historical `1.0.0` section.

No external release operation is performed by this milestone.
