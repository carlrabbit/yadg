# Milestones

| ID | Title | State | Purpose |
|---|---|---|---|
| M0001 | Initial implementation substrate | complete | Establish the .NET solution, Office-independent authoring architecture, textual-tag OOXML proof, CLI/check path, and repository validation interface. |
| M0002 | Workspace and core document authoring | complete | Establish workspace-aware authoring, explicit section/content tags, structured Markdown blocks/inlines, global stable references, and multi-template DOCX output. |
| M0003 | Structured content authoring | complete | Add lists, generated tables, figures/assets, captions, and explicit float-like table/figure placement. |
| M0004 | References and Word document structures | complete | Add template-local front matter/prototypes, numbered captions, bookmarks/SEQ/REF fields, and semantic numeric references. |
| M0005 | LibreOffice renderer and finalization | complete | Add real LibreOffice field/index finalization, finalized DOCX/PDF artifacts, runtime isolation, renderer validation, and blocking artifact-quality review. |
| M0006 | Workspace values | complete | Add source-controlled scalar workspace values in `YADG.md` and inline template `{{value:<id>}}` substitution while preserving template-owned formatting. |
| M0007 | Prepared table row population | complete | Populate template-owned Word table body rows from semantic Markdown tables while preserving template presentation. |
| M0008 | External content producers and Mermaid diagrams | complete | Add a constrained external-process content-producer boundary and inline Mermaid fenced blocks producing PNG-backed semantic figures. |
| M0009 | Microsoft Word renderer and publishing | ready | Add real Microsoft Word finalization on Windows and explicit publication of finalized DOCX to a configurable/CLI-selected delivery path. |

Milestone implementation starts from the milestone document and its explicitly listed authority.
