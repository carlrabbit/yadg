# Troubleshooting

## Find the source location

Diagnostics use workspace-relative paths. Markdown and `YADG.md` diagnostics include one-based line/column when a concrete token is available. DOCX diagnostics use a structural location such as a body paragraph or prototype because Word documents do not have useful source lines. Warnings are displayed but do not fail the command; errors do.

Run `yadg check --list` to see what YADG discovered alongside errors. The listing is a human aid, not a stable data format.

## Common fixes

- **No templates found:** put at least one prepared `.docx` directly in `YadgTemplates/`.
- **Unsupported thematic break:** remove it or set `markdown.thematicBreak: ignore`.
- **Unsupported inline code:** choose `markdown.codeInline: ignore` to keep the text plain, or `style` and bind a character style as `styles.codeInline` in template front matter.
- **Missing or unresolved reference:** check the stable ID spelling and ensure the object is selected/rendered by the template.
- **Wrong style type / unresolved style:** run `yadg inspect styles`, then select the exact ID, Word display name, or alias appropriate to the role. Names and aliases are case-insensitive exact matches; IDs take exact precedence.
- **List style lacks numbering:** Word's displayed name can differ from the internal ID, aliases can name the same style, and a gallery entry may not be a concrete serialized style. Inspect available styles; select a concrete usable list paragraph style or bind a real numbered/bulleted `{{item}}` prototype.
- **Missing output directory:** output directories are created by their owning commands; `check` does not require them.
- **Mermaid producer failure:** verify its executable, arguments, PATH, and output support. YADG does not install or sandbox producers.
- **Word render failure:** use an interactive logged-in session with activated Word and clear any modal prompt.
- **LibreOffice render failure:** install Writer and use `--renderer-path` when `soffice` is not on PATH.
- **Publish failure:** render first, ensure at least one top-level finalized DOCX exists, and set `--publish-path` or `publish.path`.

For the full workflow see [Getting started](GETTING-STARTED.md), and for supported Markdown/configuration see [Authoring](AUTHORING.md).
