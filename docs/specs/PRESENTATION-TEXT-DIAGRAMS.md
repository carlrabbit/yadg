# Built-in text-grid diagrams in PowerPoint (M0020)

## Status and precedence

Normative M0020 specialization of `TEXT-DIAGRAMS.md`, `PRESENTATION-AUTHORING.md`, and `CROSS-FORMAT-AUTHORING.md` for **built-in** character-grid diagrams rendered inside PPTX image regions. Existing DOCX diagram semantics and external Mermaid/PNG producer semantics remain in force. This file overrides older PPTX image-region text that lists only PNG/JPEG and Mermaid for this one additional source type.

## Product boundary

A box/connector diagram authored as character-grid Markdown is processed by the **existing Office-independent `TextDiagramParser` and geometry/occlusion model**, producing its existing deterministic, passive, self-contained **SVG**. The diagram is not a Mermaid producer and needs no external diagram CLI. Both DOCX and PPTX consume the shared SVG product without maintaining divergent parsers or changing the source's geometry. The PPTX-specific adapter owns packaging as an **ordinary native PowerPoint SVG picture**, not editable individual PowerPoint boxes, connectors or text runs. The picture remains movable/resizable/replaceable as one object; component-wise editing is explicitly unsupported. Editing Markdown, not the picture internals, is the canonical change path.

SVG must remain present as the vector payload in the authored PPTX and must be the intended primary representation in recent native PowerPoint and LibreOffice Impress. The implementation may include a technically required, bounded companion compatibility bitmap/preview as an **auxiliary** Open XML fallback when needed for valid Office interoperability, but a PNG-only substituted diagram, a rasterized-slide picture, or an external SVG-to-PNG producer as the normal diagram path does **not** satisfy M0020. Preserve source SVG dimensions/viewBox and scale it via contain-in-region geometry; never crop, stretch disproportionately, overrun template bounds, or change template layout/fonts. Implementer may choose compliant Office SVG relationship/fallback mechanics based on real Open XML/native application interoperability; do not introduce new author-facing configuration, external renderer dependency, or an SVG file-attachment feature.

The SVG is generated only from YADG's built-in text-diagram source. Arbitrary `.svg` images referenced by Markdown remain unsupported. The SVG is ephemeral, with no persistent workspace cache or user-visible generated source file. No post-flight generic overflow analysis is introduced.

## Presentation Markdown grammar and placement

An `image`-typed region may contain exactly one of:

- an existing local PNG/JPEG Markdown image;
- an existing identified Mermaid fence;
- an explicit identified character-grid `diagram` fence; or
- one **qualifying** opt-in auto-detected diagram fence.

Explicit form (the same grammar as DOCX):

````markdown
# Architecture {#text-image}

## Notes {#text}

Built-in source-grid drawing.

## Diagram {#image}

```diagram {#system-map caption="System map"}
+------------------+
| Service          |
| +--------------+ |
| | Worker       | |
| +--------------+ |
+------------------+
```
````

The explicit fence requires a stable ID using existing ID syntax and permits only the existing optional `caption="..."`. Its ID is unique **within the deck** across ordered source files, like existing PPTX Mermaid image IDs; this is not a DOCX global figure/reference namespace. The optional caption is image metadata/accessible description, **not** a slide-visible numbered caption. Preserve plain text within the SVG and use meaningful accessible text/alt metadata where supported; do not synthesize references or cross-slide numbering.

`diagrams.autoDetect: true` has exactly the existing DOCX candidate/heuristic semantics: unlabeled fences and exact `text` fences only; strong renderable geometry required; ambiguous/weak/nonqualifying blocks retain ordinary PPTX unsupported-code semantics. An auto-detected `text {#id caption="..."}` fence has explicit deck-scoped ID; an unlabeled detected fence may be **anonymous** within one image region and must not acquire a public synthetic ID or reference. One diagram per typed image region, and other text/content in that region remains incompatible. No detection outside fences, in a text/table region, or in raw Markdown paragraphs. Default autoDetect is false. Direct presentation authoring inherits this setting only from explicitly supplied `--config`, while workspace/selected modes read it from `YADG.md` as already specified. Neither direct CLI nor workspace schema increments its version.

## Strict, YOLO and transactions

Use the **existing** shared text-diagram parse/error semantics for invalid explicit geometry, ambiguous occlusion, missing ID and malformed attributes. These remain **fatal under strict and YOLO**; do not recover by inventing shapes, omitting content or using an arbitrary placeholder. Qualifying auto-detection uses the established conservative success/false-negative rules, rather than claiming recoveries for weak candidates. SVG authoring/packaging/interoperability failure is fatal with contextual slide/region/source diagnostic and must leave existing deck output bytes unchanged, including selected-deck and direct one-file transactions. No silently substituted PNG-only output or fallback to a different template. An explicit diagram does not require or invoke `producers.mermaid`.

## Validation and compatibility

- Focused integration tests must compare DOCX and PPTX use of identical diagram source through the shared geometry/SVG product. Cover nested boxes, connectors/arrows, unambiguous occluded edges, stable IDs, caption metadata, explicit invalid geometry, deck duplicate IDs, text- and table-region rejection, anonymous and identified auto-detection, off-by-default behavior and no-output-mutation failure.
- Verify PPTX picture XML and relationships carry **actual generated SVG content**; verify contain geometry, meaningful accessibility metadata and template-owned slide/image bounds. If companion fallback exists, verify it is auxiliary to SVG and not a silent raster-only replacement. Verify repeated image regions/slides don't reuse incorrect part IDs/content.
- Authoritative Tier 3 tests produce from **both** current application-origin PPTX fixtures on an authorized interactive Windows host: PowerPoint-origin template → YADG → real PowerPoint, and Impress-origin template → YADG → real LibreOffice Impress. The same source diagram(s) must display correctly with intact vectors/labels/lines/containment, native SVG picture semantics (not constituent PowerPoint shapes), preserved layout/master, no repair, and no clipping/overflow. Real native-application slide renders must be retained alongside per-origin receipts, version and candidate/fixture hashes; cross-opening alone is insufficient.
- Tier 4 installs **the current packed tool** into an isolated path and exercises the documented source/config/bootstrap workflow with an SVG diagram in a PPTX image region; the installed output must be accepted by the declared real native boundary. Existing DOCX diagrams/Mermaid images must not regress.
- Human artifact-quality gate is **blocking** for two distinct origin-matched PowerPoint presentation artifacts with one or more meaningful diagrams, inspected in PowerPoint and Impress respectively. Exactly one reviewer decision (HR-M0020-01) must bind their SHA-256 and Tier 3/4 receipts and inspect all slide renders for layout, vector fidelity, legibility, containment and editability-as-one-picture. Approval cannot be fabricated. No Word artifact is added to this *new* human gate, but DOCX regressions have automated coverage.

## Documentation and workspace contracts

The implementation must add a concise, example-driven **`docs/user/BOX-DIAGRAMS.md`** guide, including Unicode/ASCII characters, whitespace preservation, nested boxes, arrows/joins, occlusion, rejection examples, explicit `diagram` fences, optional `text`-fence auto-detection and DOCX/PPTX differences, plus examples validated in regression/consumer tests. Add **`docs/user/CONFIGURATION.md`** as a complete, consolidated, annotated `YADG.md` **schema-v1** user-facing reference: every supported root (`yadg`, `values`, `markdown`, `diagrams`, `references`, `producers`, `presentations`, `publish`), legal child keys/types/defaults, paths/trust/unknown-key behavior, direct `--config` subset and location semantics, and representative combined workspace example. This is documentation of current behavior, not permission to add keys or change parser semantics. Update root `README.md` links/example and relevant Markdown and presentation guides; repair contradictions such as saying diagrams are DOCX-only or source values never apply to Markdown. Add a sample checked by `yadg init` + CLI, where practical; do **not** demand that intentionally malformed unit/integration fixtures be created through bootstrap.
