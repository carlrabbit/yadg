# Authoring

YADG authoring is now documented as two explicit disciplines:

- [Markdown authoring](MARKDOWN-AUTHORING.md) — source Markdown, semantic IDs/references, tables/figures, built-in text diagrams, Mermaid, and source-facing policy.
- [DOCX template authoring](DOCX-TEMPLATE-AUTHORING.md) — visible template tags, styles, prototypes, prepared tables, captions/references, values, inspection, and Word-owned presentation.

For the end-to-end workflow, start with [Getting started](GETTING-STARTED.md).

For one-off authoring without a workspace, `yadg document check/build` accepts an explicit DOCX template and ordered Markdown sources. `document build` writes one authored, unfinalized DOCX. Both `document` and `presentation` commands accept an optional explicit `--config` and invocation-local `--yolo`; neither reads an ambient `YADG.md`. See the format guides for command examples and the direct-config trust boundary.

This path is retained so existing links to `docs/user/AUTHORING.md` continue to lead to useful current documentation.
