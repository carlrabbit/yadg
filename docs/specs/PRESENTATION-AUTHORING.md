# Template-based PowerPoint authoring (M0018)

## Ownership and boundary

YADG creates editable PPTX presentations from Markdown and ordinary prepared PPTX templates. Markdown owns slide order, titles, semantic content and logical layout/region selection. A reusable prototype slide in the PPTX owns shape positions, page/master/theme styling, static artwork, fonts and formatting examples. YADG duplicates the selected prototype in source order and fills designated regions; it does not generate CSS/HTML slides, rasterize a slide, require Office to author PPTX, or derive geometry from Markdown. This is a presentation output path, not a change to existing DOCX semantics.

## Slide Markdown grammar

- Every ATX level-1 heading begins a slide. Its visible heading text supplies the slide title. `{#layout}` on that heading selects a logical layout. With no selector, use `content`.
- Each level-2 heading selects a region by its `{#region}` ID; its visible text is for Markdown readability only and is not emitted. Region selectors are required on level-2 headings. Both layout and region IDs intentionally repeat across slides; they are *selectors* in presentation mode, not document-unique reference IDs.
- Level-3 through level-6 headings are rendered as content headings, with styling relative to the region's prototype examples. They never select regions. Do not silently rebase or reinterpret these headings as slide/region delimiters.
- For a layout with one non-title region, content directly after the level-1 heading and before the next level-1 heading goes into that region; no level-2 heading is necessary. For multi-region layouts, all content must belong to a selected level-2 region. For single-region layouts an explicit level-2 selector is also legal, but content may not be supplied both directly and via that selector.
- Only ordinary Markdown heading IDs (Markdig generic attributes) are used: no new layout syntax. Slide selector IDs do not create unique cross-reference anchors; cross-slide reference/citation semantics are out of scope.
- A layout may be instantiated any number of times. A selected region may occur at most once within one slide. Region ordering in Markdown does not dictate physical arrangement.
- Reject undefined layouts, undefined/missing/duplicate regions, conflicting direct and region-scoped content, unsupported block types in a typed region, unexpected content before the first slide, and empty/invalid required payloads with source path and slide/region diagnostics. No silent truncation.

Example:

```markdown
# Architecture {#text-image}

## Explanation {#text}

### Backend
- Markdig
- Open XML

## Diagram {#image}

![System overview](images/overview.png)

# Results {#table}

| Metric | Value |
|---|---:|
| Throughput | 42 |
```

## Standard layout catalog

The catalog defines logical region IDs and types only; template geometry is unconstrained. A template need implement only the layouts used by its sources; no implicit substitution of a missing layout by another. The slide title is a separate `title` slot whenever shown below. `title` and `section` layouts may style the heading differently. `section` has no content region.

| Layout selector | Content regions (`id:type`) |
|---|---|
| `title` | `content:text` |
| `section` | none |
| `content` | `content:text` |
| `two-columns` | `left:text`, `right:text` |
| `three-columns` | `left:text`, `center:text`, `right:text` |
| `two-rows` | `top:text`, `bottom:text` |
| `three-rows` | `top:text`, `middle:text`, `bottom:text` |
| `grid-2x2` | `top-left:text`, `top-right:text`, `bottom-left:text`, `bottom-right:text` |
| `image` | `image:image` |
| `text-image` | `text:text`, `image:image` |
| `image-text` | `image:image`, `text:text` |
| `image-above-text` | `image:image`, `text:text` |
| `text-above-image` | `text:text`, `image:image` |
| `table` | `table:table` |
| `text-table` | `text:text`, `table:table` |
| `table-text` | `table:table`, `text:text` |

Distinct mixed-layout names explicitly communicate *intended* relative placement, while exact geometry and styling remain template-owned. No standardized layout has more than four non-title regions; mixed image/table slides use at most two. Custom layouts and arbitrary grid composition are deferred.

## Typed region payloads

- `text`: paragraphs, emphasis/strong, inline code according to configured Markdown policy, links as displayed text/link when supported, heading levels >=3, ordered/unordered lists with nesting, and line breaks. Unsupported structures, including raw HTML, figures and tables inside text regions, produce diagnostics rather than content loss. Text shape dimensions and existing fit settings are preserved; no automatic font-size/layout adjustment by YADG.
- `image`: exactly one Markdown image block with local workspace-relative PNG/JPEG path, alt text and optional title; no other non-whitespace blocks. Reuse existing workspace asset path/validation rules. Preserve aspect ratio using contain-within-region (letterbox if needed); never distort, crop or expand bounds silently. The picture itself is an editable ordinary PowerPoint picture shape.
- `table`: exactly one Markdown pipe table with header and body rows; no surrounding prose. Create/extend a native editable PowerPoint table using a template-defined header/body-cell example; do not rasterize. Preserve header/body cell formatting, column widths/proportions and table position, fitting row count only when feasible. Any incompatibility or inability to fit according to this milestone's deterministic *structural* checks must be fatal; do not claim visual-fit validation.
- All slots are independently typed; never mix image/table/text payload types inside one region. Text-only layouts may contain multiple headings/paragraphs/lists in their text regions.

## Authoring the PowerPoint template

- A prototype is an ordinary editable slide. It carries one visible plain-text marker `{{slide:<layout-id>}}` in a dedicated shape. The marker is removed from generated slides; each layout ID must identify exactly one prototype. All other static elements are preserved when duplicated.
- A title-bearing prototype contains exactly one plain `{{title}}` marker in its title text shape. A region contains exactly one plain marker `{{<region-id>}}` (for example `{{left}}`, `{{image}}`, `{{table}}`). Marker text must be literal/unambiguous, not embedded in prose. Markers must not be shared across overlapping text containers or hidden in shape metadata.
- Text regions are ordinary text shapes with the marker and representative paragraph/list/heading formatting examples in the prototype; formatting samples are presentation resources, not emitted literal content. They must enable YADG to distinguish plain paragraph, content-heading and list role appearances. If a required role has no example, report a precise template incompatibility rather than inventing a font/style.
- Image regions are ordinary bounded text shapes bearing `{{image}}`. Replace the marked shape with a native picture occupying that region's bounds; preserve other slide elements. No crop/shape-masking model in M0018.
- Table regions use an editable native PowerPoint table with `{{table}}` as the sole marker in the top-left cell; at least one representative header row and one representative body row are required. The template provides table geometry, column count, column widths, and representative row/cell properties. Input column count must equal prototype column count. Append body rows by duplicating the body-row example, while keeping output editable. Unsupported merges/nesting are template errors; output that cannot be represented without reflow must fail structurally.
- The prototype slide itself is not included in generated presentations unless separately instantiated from Markdown. Prototype discovery must not rely on slide number. Keep original slide/master/theme relationships, media and static shapes when cloning; generated slide order follows Markdown order. Unselected prototype slides do not appear as output content. Explicitly reject ambiguous or malformed markers, and do not accidentally replace static text elsewhere.
- Template-authoring remains possible in PowerPoint or LibreOffice Impress; no mandatory shape names, hidden metadata, content controls, or sidecar layout definitions. Simple visible tags are the contract.

## Application-origin compatibility fixtures and canonical review deck (M0018 correction)

This section governs **test and review fixtures**, not the production format of user templates. It adopts the existing M0010 application-produced-fixture rule for PowerPoint.

- Maintain **exactly two canonical committed compatibility PPTX templates** for M0018: one Microsoft PowerPoint-origin and one LibreOffice Impress-origin. They are separate test inputs, not three competing review templates. No third canonical M0018 template is permitted. Temporary or narrowly scoped negative-case fixtures may exist independently but must be clearly identified as such and may not be presented as canonical application-origin fixtures.
- The PowerPoint-origin template must start as a **new presentation created through the real desktop Microsoft PowerPoint document/application object model** on an interactive Windows host with activated PowerPoint. The Impress-origin template must start as a **new presentation created through real LibreOffice Impress UNO**. Create the required slide and text/table structures through that application's document model, then save as PPTX. Constructing OOXML from scratch and subsequently opening/saving it in the application does **not** establish provenance.
- In each application, retain its genuine default/native master, theme, layout relationships and representative title/body/list/table formatting. Starting from a blank file and manually reproducing unstyled shapes, or stripping inherited styles and replacing them with arbitrary hardcoded defaults, does not satisfy this condition. Prototype markers and examples may be added after creating the native presentation. These are synthetic, redistribution-safe fixtures and must not contain corporate or confidential material.
- **Required fixture coverage, for each origin:** usable YADG prototypes for `content` (ordinary paragraphs and content heading), `two-columns` (independent text slots with representative lists), `text-image` (text and one native picture region), and `table` (native editable table with header and body examples). These four contracts form the fixture compatibility sample, **not** a requirement that all sixteen standard layouts be present in every fixture. Other standard layouts remain covered by focused tests as appropriate. Every representative role must inherit or deliberately retain real application formatting.
- Commit provenance for **each** canonical template: repository path, originating application and exact version, OS/platform, UTC creation/save date, creation method (PowerPoint automation or Impress UNO), SHA-256 of the committed PPTX, synthetic/redistribution-safe declaration, and a statement that the file originated in the application rather than an Open XML generation/round-trip. Tests must verify committed bytes against the recorded hashes and use copies, not rewrite canonical fixture bytes.
- Tier 3 compatibility must generate a representative deck from **each** origin and open it in its respective originating application: PowerPoint-origin output in real desktop PowerPoint; Impress-origin output in real Impress. Check no repair prompt, native editable text/picture/table content, correct slide order, preservation of master/theme/static elements and visible formatting inherited from the prototype. Also check opening the designated review deck in Impress. One renderer cannot substitute for the other origin-specific boundary.
- There must be **exactly one named human-review candidate**, an Impress-origin generated PPTX covering the four representative prototype contracts, at `artifacts/review/evidence/M0018/review-sample.pptx`. A review manifest must bind its path/SHA-256, source/template fixture identity and SHA-256, repository revision, application/runtime versions and the corresponding Tier 3 evidence. Other generated PPTX files are integration-test artifacts only and must not be presented as additional review candidates. The reviewer is instructed to open this one file in LibreOffice Impress and examine fidelity, readable formatting, slide order, absence of leaked markers, native text/image/table editability, static content/master preservation and absence of repair warnings. Overflow may be noticed in human review but is not automatically measured.
- On correcting existing fixtures, identify references to superseded sample files and remove or relabel only obsolete artifacts after their consumers have been updated. Do not indiscriminately delete negative-test fixtures or historical evidence. Pending review evidence must be regenerated against the corrected inputs and the current implementation; do not carry forward an old approval or fabricate one.
- **Locus and blocked semantics:** PowerPoint-origin creation and open validation require activated desktop PowerPoint in an interactive Windows session; Impress-origin creation and open validation require installed LibreOffice Impress on an authorized interactive Windows host (isolated user profile for validation). Existing authorized installations may be used. Machine-wide provisioning, new CI runners/workflows and unapproved privileges are not authorized. If a runtime is absent, its corresponding fixture or validation obligation stays **BLOCKED**; never label a synthetic/round-tripped substitute as application-produced or pass the other renderer's evidence as equivalent.
- Record a verifiable success/failure receipt for both origin-specific runs including application/version, template/output hashes, input/source identity, repository revision, output inspection/reopen result and cleanup status. The reviewer decision remains independently blocking; automated inspection or an implementer's own confirmation does not replace it.

## Workspace, CLI and output

### Public `YADG.md` schema-v1 declaration (M0018 correction)

`presentations` is an **optional root sequence** of named deck declarations. Its full schema is:

```yaml
---
yadg:
  version: 1
presentations:
  - id: project-overview
    template: YadgTemplates/corporate.pptx
    sources:
      - slides/overview.md
      - slides/results.md
    output: project-overview.pptx
  - id: technical-detail
    template: YadgTemplates/technical.pptx
    sources:
      - slides/details.md
    output: technical-detail.pptx
---
```

`presentations` is a non-empty YAML sequence when present. Each entry is a mapping with exactly four required keys: `id`, `template`, `sources`, `output`. Unknown keys, duplicate mapping keys, non-mappings, nulls, and implicit scalar coercion are errors. `id` is a case-sensitive non-empty stable declaration identifier matching `[A-Za-z][A-Za-z0-9_-]*`; IDs must be unique (ordinal/case-sensitive). `template` and `output` are non-empty YAML string scalars; `sources` is a non-empty ordered sequence of non-empty YAML string scalars. A deck may have multiple sources, and there may be multiple deck declarations; **there is no M0018 limit of one deck**. The singular root `presentation` is **not** supported; reject it with a diagnostic pointing to `presentations`. Do not silently accept both forms or auto-migrate configuration. The `yadg.version` remains `1` and the existing root YAML restrictions still apply. Absent `presentations` means no configured PPTX work and leaves existing workspaces unchanged; nothing is inferred from discovered files.

For each declaration, `template` names an existing valid `.pptx` under `YadgTemplates/`, resolved relative to the workspace root with lexical containment enforced. Each `sources` entry names an existing `.md` file relative to the workspace root, normalized, lexically within the workspace, outside input-template/generated directories and excluded dot directories; sources inside **one** deck must be distinct, with no wildcards, directories or URLs. **Sharing the same source file across different declarations is allowed**, as is reusing a template. For each deck, concatenate slides from its sources in sequence order; each file starts a fresh slide sequence and may not begin with stray content. Markdown images resolve relative to their owning source using existing figure lexical-containment rules. `output` is a plain `.pptx` basename (no path components, reserved/dot names or invalid filename characters) and must be distinct across all declarations using case-insensitive filename comparison to avoid collisions on Windows. The destination is `YadgPresentations/<output>`.

The **union** of all declared presentation source paths is excluded from existing recursive DOCX Markdown discovery, even if a file participates in several decks. Other discoverable Markdown sources retain current DOCX meaning. Presentation-only, DOCX-only, and mixed workspaces remain valid. Repeated layout/region selectors obey presentation-scoped semantics, never DOCX global semantic-ID uniqueness. Report declaration ID and source/slide/region where applicable.

`yadg check` validates **every** configured presentation and independently applicable DOCX inputs. `yadg build` generates **every** configured deck and applicable DOCX outputs. For the PPTX result set, validate/stage **all** decks before committing: success replaces the complete top-level `YadgPresentations/*.pptx` set with exactly the declared outputs (clearing stale outputs), while any PPTX failure preserves the complete previous successful PPTX set, not merely the failing deck. The PPTX result set is independent of existing DOCX/PDF commit sets; do not claim cross-format atomicity. `render` and `publish` remain DOCX-only. `init` need not create any PPTX starter. `--yolo` may not silently discard malformed declarations, invalid slides or a failing deck.

### Direct commands without `YADG.md`

Support a second, **configuration-free** invocation path. These are public commands in M0018, not merely hypothetical examples:

```powershell
yadg presentation check --template ./templates/corporate.pptx --source ./slides/part1.md --source ./slides/part2.md
yadg presentation build --template ./templates/corporate.pptx --source ./slides/part1.md --source ./slides/part2.md --output ./out/deck.pptx
```

- Both commands require **exactly one** `--template <file.pptx>` and **one or more** repeatable, ordered `--source <file.md>` arguments. `presentation build` additionally requires exactly one `--output <file.pptx>`; `presentation check` must reject `--output`. Unrecognized/duplicate single-valued options fail visibly; no implicit file discovery or template selection.
- Direct mode **does not read, require, create, or validate `YADG.md`**, and does not execute DOCX authoring, producer commands, render, or publish. A malformed `YADG.md` in the current directory must not affect it. No workspace marker or `YadgTemplates` directory is required. Path arguments are resolved against the caller's current working directory; source-relative image references resolve relative to their Markdown file. Files must be local; no URLs, globs or directories. Existing local-file asset lexical-containment rules apply with the effective root equal to the current working directory; explicit source and template paths may be located outside the current directory by the ordinary OS path rules, but images must stay lexically inside this effective root. The direct command must not implicitly adopt a nearby `YADG.md` or its `markdown`, `values`, or `producers` settings: use the standard default presentation Markdown policy. Unsupported external-producer-dependent content fails explicitly rather than silently executing a configured producer.
- `presentation build` writes **only** the explicitly named output file (creating its parent directory when valid); it must not modify `YadgPresentations/` or other decks. Validate and generate into a temporary sibling file and replace the target only on success; failure preserves a prior valid target. `presentation check` does not write PPTX output. Configured mode retains its separate complete-result-set behavior.
- Existing `yadg check`/`yadg build` behavior continues to use `YADG.md` when applicable; the `presentation` subcommands are explicitly the standalone path. No new `--presentation <id>` selection flags on existing commands are required in M0018. To produce a subset without loading workspace configuration, use the direct command.

`YadgPresentations/` is a generated-result directory, excluded from recursive Markdown discovery, and users must not store other persistent files there. The direct output path is not constrained to this generated directory because the user specified it explicitly.

### Scope boundary

Output is native editable PPTX with text, pictures and tables; do not depend on an Office runtime for generation. No external image links. PPTX inspection/real Impress reopen is validation, not part of ordinary build. No static unmarked template slides are emitted.

## Validation and out-of-scope rendering guarantees

- Pre-flight `check` must deterministically validate selectors, region completeness/types, markers, structure, assets and table column count.
- Post-flight layout/overflow detection (including UNO, PowerPoint COM and image/table/text rendered bounds), reflow, visual AutoFit measurement and post-export PDF comparisons are deliberately deferred. Preserve template AutoFit flags, but do not assert fit. The generated PPTX may have visually overflowing content; this milestone does not diagnose such overflow.
- Automated integration evidence includes inspecting PPTX package structure and reopening the generated deck with a real LibreOffice Impress runtime on an authorized Windows host. Human artifact-quality review checks real prototype fidelity, visual readability and native editability; obvious overflow is a review finding, not evidence of an implemented automatic overflow checker.
