# Agent Instructions

## Default implementation path

For implementation work:

1. Read `docs/ENGINEERING.md`.
2. Read the active milestone in `docs/milestones/`.
3. Read only the project-authority documents explicitly listed by that milestone.
4. Inspect relevant source and test files.
5. Decompose the milestone into bounded work packages and maintain `.execution/<milestone-id>.md` while executing.

Do not require the external guide repository to implement YADG.

## Coordination metadata

During ordinary implementation, ignore unless explicitly in scope:

- `.guide-profile.json`;
- `.guide-sync/`.

These files support planning and documentation synchronization. They are not ordinary implementation authority.

## Scope discipline

Do not:

- invent project semantics that are absent from the milestone or project authority;
- treat generated DOCX/PDF artifacts as source authority;
- move Microsoft Word/Office Interop dependencies into the Office-independent authoring core;
- replace visible text template tags with Word content controls unless project authority is explicitly changed;
- introduce a generic test pyramid requirement; YADG uses integration-first validation where boundary behavior materially determines correctness;
- perform broad documentation synchronization unless explicitly requested.

If implementation requires a new material decision about architecture, semantics, compatibility, scope, acceptance criteria, or validation policy, return that decision to planning instead of silently choosing it.
