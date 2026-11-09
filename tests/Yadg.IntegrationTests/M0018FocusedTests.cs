using Yadg.Core;
using Yadg.Presentation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OpenXmlPresentationDocument = DocumentFormat.OpenXml.Packaging.PresentationDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace Yadg.IntegrationTests;

public sealed class M0018FocusedTests
{
    [Fact]
    public void Canonical_native_office_fixtures_bind_four_layouts_to_their_committed_hashes()
    {
        var repository = FindRepositoryRoot();
        var fixtureDirectory = Path.Combine(repository, "tests", "fixtures", "m0018");
        Assert.Equal(2, Directory.GetFiles(fixtureDirectory, "*-origin.pptx", SearchOption.TopDirectoryOnly).Length);
        var expectedLayouts = new[] { "content", "two-columns", "text-image", "table" };

        foreach (var id in new[] { "powerpoint-origin", "impress-origin" })
        {
            var path = Path.Combine(fixtureDirectory, id + ".pptx");
            using var provenance = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(fixtureDirectory, id + ".provenance.json")));
            var record = provenance.RootElement;
            var digest = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path)));
            Assert.Equal(digest, record.GetProperty("sha256").GetString());
            Assert.Equal(id, record.GetProperty("id").GetString());
            Assert.True(record.GetProperty("createdInApplication").GetBoolean());
            Assert.True(record.GetProperty("notOpenXmlGeneratedOrRoundTripped").GetBoolean());
            Assert.True(record.GetProperty("retainedNativeMasterThemeLayouts").GetBoolean());
            Assert.True(record.GetProperty("syntheticAndRedistributionSafe").GetBoolean());
            Assert.False(string.IsNullOrWhiteSpace(record.GetProperty("applicationVersion").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(record.GetProperty("platform").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(record.GetProperty("creationMethod").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(record.GetProperty("createdAtUtc").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(record.GetProperty("saveDateUtc").GetString()));
            Assert.Equal(expectedLayouts, record.GetProperty("layouts").EnumerateArray().Select(value => value.GetString()));

            using var document = OpenXmlPresentationDocument.Open(path, false);
            var presentationPart = document.PresentationPart!;
            var slides = presentationPart.Presentation.SlideIdList!.Elements<P.SlideId>()
                .Select(slideId => (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!)).ToArray();
            Assert.Equal(4, slides.Length);
            Assert.Single(presentationPart.SlideMasterParts);
            Assert.NotNull(presentationPart.SlideMasterParts.Single().ThemePart);
            Assert.Equal(expectedLayouts, slides.Select(slide => Assert.Single(slide.Slide.Descendants<A.Text>(), text => text.Text.StartsWith("{{slide:", StringComparison.Ordinal)).Text[8..^2]));
            Assert.All(slides, slide => Assert.Contains(slide.Slide.Descendants<P.Shape>(), shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Native template footer"));
            Assert.Contains(slides[0].Slide.Descendants<P.Shape>(), shape => shape.TextBody?.InnerText.Contains("{{content}}", StringComparison.Ordinal) == true);
            Assert.Equal(2, slides[1].Slide.Descendants<P.Shape>().Count(shape => shape.TextBody?.InnerText.Contains("{{left}}", StringComparison.Ordinal) == true || shape.TextBody?.InnerText.Contains("{{right}}", StringComparison.Ordinal) == true));
            Assert.Contains(slides[2].Slide.Descendants<P.Shape>(), shape => shape.TextBody?.InnerText.Contains("{{text}}", StringComparison.Ordinal) == true);
            Assert.Contains(slides[2].Slide.Descendants<P.Shape>(), shape => shape.TextBody?.InnerText.Contains("{{image}}", StringComparison.Ordinal) == true);
            var nativeTable = Assert.Single(slides[3].Slide.Descendants<A.Table>());
            Assert.Equal("{{table}}", nativeTable.Descendants<A.TableCell>().First().InnerText);
            Assert.Equal(2, nativeTable.Elements<A.TableRow>().Count());
        }
    }

    [Fact]
    public void Workspace_presentations_are_an_ordered_sequence_and_allow_cross_deck_sharing()
    {
        var source = "---\nyadg:\n  version: 1\npresentations:\n  - id: overview\n    template: YadgTemplates/main.pptx\n    sources:\n      - slides/a.md\n      - slides/b.md\n    output: overview.pptx\n  - id: details\n    template: YadgTemplates/main.pptx\n    sources:\n      - slides/a.md\n    output: details.pptx\n---\n";
        var diagnostics = new List<Diagnostic>();

        var values = WorkspaceValuesParser.Parse(source, "YADG.md", diagnostics);

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.IsError);
        Assert.Equal(new[] { "overview", "details" }, values.Presentations.Select(deck => deck.Id));
        Assert.Equal(new[] { "slides/a.md", "slides/b.md" }, values.Presentations[0].Sources);
        Assert.Equal("slides/a.md", Assert.Single(values.Presentations[1].Sources));
    }

    [Fact]
    public void Workspace_presentation_configuration_rejects_singular_duplicate_and_invalid_output_declarations()
    {
        var diagnostics = new List<Diagnostic>();
        WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\npresentation:\n  id: old\n---\n", "YADG.md", diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "YADG-PRESENTATION-CONFIG-001");

        diagnostics.Clear();
        WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\npresentations:\n  - id: deck\n    template: t.pptx\n    sources:\n      - a.md\n    output: same.pptx\n  - id: other\n    template: t.pptx\n    sources:\n      - b.md\n    output: SAME.pptx\n---\n", "YADG.md", diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "YADG-PRESENTATION-CONFIG-008");

        diagnostics.Clear();
        WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\npresentations:\n  - id: deck\n    template: t.pptx\n    sources:\n      - a.md\n    output: CON.pptx\n---\n", "YADG.md", diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "YADG-PRESENTATION-CONFIG-007");
    }

    [Fact]
    public void Workspace_sources_declared_by_multiple_decks_are_excluded_from_docx_discovery()
    {
        using var workspace = WorkspaceFixture.Create();
        Directory.CreateDirectory(Path.Combine(workspace.Root, "YadgTemplates"));
        Directory.CreateDirectory(Path.Combine(workspace.Root, "slides"));
        File.WriteAllText(Path.Combine(workspace.Root, "YadgTemplates", "deck.pptx"), "pptx");
        File.WriteAllText(Path.Combine(workspace.Root, "YadgTemplates", "report.docx"), "docx");
        File.WriteAllText(Path.Combine(workspace.Root, "slides", "shared.md"), "# Presentation\n\nSlide body.");
        File.WriteAllText(Path.Combine(workspace.Root, "docx.md"), "# Document {#document}\n\nBody.");
        File.WriteAllText(Path.Combine(workspace.Root, "YADG.md"), "---\nyadg:\n  version: 1\npresentations:\n  - id: first\n    template: YadgTemplates/deck.pptx\n    sources:\n      - slides/shared.md\n    output: first.pptx\n  - id: second\n    template: YadgTemplates/deck.pptx\n    sources:\n      - slides/shared.md\n    output: second.pptx\n---\n");

        var workspaceData = WorkspaceLoader.Load(workspace.Root);

        Assert.Equal(new[] { Path.Combine(workspace.Root, "docx.md") }, workspaceData.Sources);
        Assert.Equal(2, workspaceData.Presentations.Count);
        Assert.DoesNotContain(workspaceData.Diagnostics, diagnostic => diagnostic.Code is "YADG-WS-003" or "YADG-WS-005");
    }

    [Fact]
    public void Presentation_parser_recognizes_all_layouts_repeated_layouts_and_ordered_slides()
    {
        var parsed = PresentationMarkdownParser.Parse(new[]
        {
            new PresentationSource("one.md", "# Intro\n\nOne.\n\n# Split {#two-columns}\n\n## Left {#left}\n\nA.\n\n## Right {#right}\n\nB."),
            new PresentationSource("two.md", "# Again {#two-columns}\n\n## Left {#left}\n\nC.\n\n## Right {#right}\n\nD.")
        });

        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        Assert.Equal(16, PresentationLayouts.All.Count);
        Assert.Equal(new[] { "Intro", "Split", "Again" }, parsed.Slides.Select(slide => slide.Title));
        Assert.Equal("content", parsed.Slides[0].Layout);
        Assert.Equal(new[] { "left", "right" }, parsed.Slides[1].Regions.Select(region => region.Id));
        Assert.Equal("two-columns", parsed.Slides[2].Layout);
    }

    [Fact]
    public void Presentation_parser_validates_the_complete_standard_layout_region_contracts()
    {
        static (string Id, PresentationRegionType Type) Text(string id) => (id, PresentationRegionType.Text);
        static (string Id, PresentationRegionType Type) Image(string id) => (id, PresentationRegionType.Image);
        static (string Id, PresentationRegionType Type) Table(string id) => (id, PresentationRegionType.Table);
        var expected = new (string Id, (string Id, PresentationRegionType Type)[] Regions)[]
        {
            ("title", new[] { Text("content") }), ("section", Array.Empty<(string, PresentationRegionType)>()), ("content", new[] { Text("content") }),
            ("two-columns", new[] { Text("left"), Text("right") }), ("three-columns", new[] { Text("left"), Text("center"), Text("right") }),
            ("two-rows", new[] { Text("top"), Text("bottom") }), ("three-rows", new[] { Text("top"), Text("middle"), Text("bottom") }),
            ("grid-2x2", new[] { Text("top-left"), Text("top-right"), Text("bottom-left"), Text("bottom-right") }),
            ("image", new[] { Image("image") }), ("text-image", new[] { Text("text"), Image("image") }),
            ("image-text", new[] { Image("image"), Text("text") }), ("image-above-text", new[] { Image("image"), Text("text") }),
            ("text-above-image", new[] { Text("text"), Image("image") }), ("table", new[] { Table("table") }),
            ("text-table", new[] { Text("text"), Table("table") }), ("table-text", new[] { Table("table"), Text("text") })
        };
        var expectedIds = expected.Select(layout => layout.Id).ToArray();
        Assert.Equal(expectedIds.Order(StringComparer.Ordinal), PresentationLayouts.All.Keys.Order(StringComparer.Ordinal));
        foreach (var layout in expected)
            Assert.Equal(layout.Regions, PresentationLayouts.All[layout.Id].Regions.Select(pair => (pair.Key, pair.Value)));

        var markdown = new System.Text.StringBuilder();
        foreach (var layout in expected)
        {
            markdown.AppendLine($"# {layout.Id} {{#{layout.Id}}}");
            foreach (var (regionId, type) in layout.Regions)
            {
                markdown.AppendLine();
                markdown.AppendLine($"## {regionId} {{#{regionId}}}");
                markdown.AppendLine();
                markdown.AppendLine(type switch
                {
                    PresentationRegionType.Text => "region text",
                    PresentationRegionType.Image => "![sample](sample.png)",
                    PresentationRegionType.Table => "| H |\n|---|\n| V |",
                    _ => throw new InvalidOperationException()
                });
            }
            markdown.AppendLine();
        }
        var parsed = PresentationMarkdownParser.Parse(new[] { new PresentationSource("layouts.md", markdown.ToString()) });
        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        Assert.Equal(expected.Length, parsed.Slides.Count);
        for (var index = 0; index < expected.Length; index++)
            Assert.Equal(expected[index].Regions, parsed.Slides[index].Regions.Select(region => (region.Id, region.Type)));
    }

    [Fact]
    public void Presentation_parser_reports_missing_duplicate_and_mistyped_regions()
    {
        var parsed = PresentationMarkdownParser.Parse(new[]
        {
            new PresentationSource("bad.md", "# Bad {#text-image}\n\n## Text {#text}\n\n![not text](pic.png)\n\n## Again {#text}\n\nDuplicate.")
        });

        Assert.False(parsed.IsValid);
        Assert.Contains(parsed.Diagnostics, diagnostic => diagnostic.Code == "YADG-PPT-MD-010");
        Assert.Contains(parsed.Diagnostics, diagnostic => diagnostic.Code == "YADG-PPT-MD-005" && diagnostic.Message.Contains("image", StringComparison.Ordinal));
        Assert.Contains(parsed.Diagnostics, diagnostic => diagnostic.Code == "YADG-PPT-MD-017");
    }

    [Fact]
    public void Presentation_preflight_rejects_malformed_source_selectors_and_payloads_without_mutating_output()
    {
        using var fixture = PresentationFixture.CreateRich();
        var sourcePath = Path.Combine(fixture.Root, "invalid.md");
        var outputPath = Path.Combine(fixture.Root, "preserved.pptx");
        var sentinel = new byte[] { 9, 8, 7, 6 };
        File.WriteAllBytes(outputPath, sentinel);
        var invalidCases = new (string Markdown, string Diagnostic)[]
        {
            ("# Unknown {#not-a-layout}\n\nBody.", "YADG-PPT-MD-002"),
            ("# Multiple {#two-columns}\n\nDirect body.", "YADG-PPT-MD-006"),
            ("# Multiple {#two-columns}\n\n## Left {#left}\n\nLeft only.", "YADG-PPT-MD-005"),
            ("# Single {#content}\n\n## Region without selector\n\nText.", "YADG-PPT-MD-009"),
            ("# Single {#content}\n\n## First {#content}\n\nOne.\n\n## Second {#content}\n\nTwo.", "YADG-PPT-MD-010"),
            ("# Single {#content}\n\n## Unknown {#other}\n\nText.", "YADG-PPT-MD-007"),
            ("# Mixed {#text-image}\n\n## Text {#text}\n\n![Image in text](image.png)\n\n## Image {#image}\n\n![First](first.png)\n\n![Second](second.png)", "YADG-PPT-MD-017"),
            ("# Data {#table}\n\n| Header |\n|---|", "YADG-PPT-MD-013"),
            ("# Mixed {#text-image}\n\n## Text {#text}\n\n| Header |\n|---|\n| Body |\n\n## Image {#image}\n\n![Image](image.png)", "YADG-PPT-MD-017")
        };

        foreach (var (markdown, expectedCode) in invalidCases)
        {
            File.WriteAllText(sourcePath, markdown);
            var diagnostics = PresentationAuthoring.Build(fixture.Template, new[] { new PresentationSource(sourcePath, markdown) }, fixture.Root, outputPath);
            var matchingDiagnostics = diagnostics.Where(item => item.Code == expectedCode).ToArray();
            Assert.NotEmpty(matchingDiagnostics);
            Assert.All(matchingDiagnostics, diagnostic => Assert.Contains("invalid.md:", diagnostic.Location, StringComparison.Ordinal));
            Assert.Equal(sentinel, File.ReadAllBytes(outputPath));
        }
    }

    [Fact]
    public void Presentation_parser_accepts_typed_image_and_table_regions()
    {
        var parsed = PresentationMarkdownParser.Parse(new[]
        {
            new PresentationSource("deck.md", "# Overview {#text-image}\n\n## Copy {#text}\n\n**Hello** and `code`.\n\n## Picture {#image}\n\n![diagram](images/diagram.png)\n\n# Data {#table}\n\n| A | B |\n|---|---|\n| 1 | 2 |")
        });

        Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.Code == "YADG-PPT-MD-017");
        Assert.IsType<PresentationImage>(Assert.Single(parsed.Slides[0].Regions.Single(region => region.Id == "image").Blocks));
        Assert.IsType<PresentationTable>(Assert.Single(parsed.Slides[1].Regions.Single(region => region.Id == "table").Blocks));
    }

    [Fact]
    public void Presentation_authoring_clones_selected_prototype_in_source_order_and_removes_controls()
    {
        using var fixture = PresentationFixture.Create();
        var sources = new[] { new PresentationSource("deck.md", "# First\n\nFirst body.\n\n# Second\n\nSecond body.") };
        var output = Path.Combine(fixture.Root, "generated.pptx");
        using (var templatePackage = OpenXmlPresentationDocument.Open(fixture.Template, false))
        {
            var templateSlide = Assert.Single(templatePackage.PresentationPart!.SlideParts);
            Assert.Contains(templateSlide.Slide.Descendants<P.Shape>(), shape => shape.TextBody?.InnerText == "{{slide:content}}");
        }

        var diagnostics = PresentationAuthoring.Build(fixture.Template, sources, fixture.Root, output);

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.IsError);
        using var result = OpenXmlPresentationDocument.Open(output, false);
        var slides = result.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(2, result.PresentationPart.Presentation.SlideIdList!.Elements<P.SlideId>().Count());
        Assert.Equal(2, result.PresentationPart.SlideParts.Count());
        var emitted = result.PresentationPart.Presentation.SlideIdList.Elements<P.SlideId>().Select(slideId => result.PresentationPart.GetPartById(slideId.RelationshipId!.Value!)).Cast<SlidePart>().ToArray();
        var titles = emitted.Select(slide => slide.Slide.Descendants<P.Shape>().Select(shape => shape.TextBody?.InnerText).First(text => text is "First" or "Second"));
        Assert.Equal(new[] { "First", "Second" }, titles);
        Assert.Contains(emitted[0].Slide.Descendants<A.Text>(), text => text.Text == "First body.");
        Assert.All(emitted, slide => Assert.DoesNotContain(slide.Slide.Descendants<A.Text>(), text => text.Text.Contains("{{", StringComparison.Ordinal)));
        Assert.Contains(slides, slide => slide.Slide.Descendants<P.Shape>().Any(shape => shape.TextBody?.InnerText == "static template element"));
    }

    [Fact]
    public void Presentation_authoring_emits_native_text_picture_and_table_objects()
    {
        using var fixture = PresentationFixture.CreateRich();
        var exportPath = Environment.GetEnvironmentVariable("YADG_M0018_TEST_TEMPLATE_EXPORT");
        if (!string.IsNullOrWhiteSpace(exportPath))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(exportPath))!);
            File.Copy(fixture.Template, Path.GetFullPath(exportPath), true);
        }
        var imageDirectory = Path.Combine(fixture.Root, "images");
        Directory.CreateDirectory(imageDirectory);
        File.WriteAllBytes(Path.Combine(imageDirectory, "pixel.png"), Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
        var source = Path.Combine(fixture.Root, "slides.md");
        File.WriteAllText(source, "# Mixed {#text-image}\n\n## Copy {#text}\n\nEditable text.\n\n### Insight\n\nA heading example.\n\n- Parent **bold**\n  - Child *italic*\n\n## Picture {#image}\n\n![Pixel](images/pixel.png)\n\n# Data {#table}\n\n| Name | Value |\n|---|---|\n| Alpha | 1 |\n| Beta | 2 |");
        var output = Path.Combine(fixture.Root, "generated.pptx");

        var diagnostics = PresentationAuthoring.Build(fixture.Template, new[] { new PresentationSource(source, File.ReadAllText(source)) }, fixture.Root, output);

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.IsError);
        using var document = OpenXmlPresentationDocument.Open(output, false);
        Assert.DoesNotContain(new OpenXmlValidator().Validate(document), error => error.Part is SlidePart);
        var slides = document.PresentationPart!.Presentation.SlideIdList!.Elements<P.SlideId>()
            .Select(slideId => (SlidePart)document.PresentationPart.GetPartById(slideId.RelationshipId!.Value!)).ToArray();
        Assert.Equal(2, slides.Length);
        Assert.Single(document.PresentationPart.SlideMasterParts);
        Assert.All(slides, slide => Assert.Single(slide.GetPartsOfType<SlideLayoutPart>()));
        Assert.Contains(slides[0].Slide.Descendants<A.Text>(), text => text.Text == "Editable text.");
        Assert.Contains(slides[0].Slide.Descendants<A.Text>(), text => text.Text == "Insight");
        Assert.DoesNotContain(slides[0].Slide.Descendants<A.Text>(), text => text.Text is "Copy" or "Picture" or "Data");
        var textShape = slides[0].Slide.Descendants<P.Shape>().Single(shape => shape.TextBody?.InnerText.Contains("Editable text.", StringComparison.Ordinal) == true);
        var generatedParagraphs = textShape.TextBody!.Elements<A.Paragraph>().ToArray();
        Assert.Contains(generatedParagraphs, paragraph => paragraph.InnerText == "Insight" && paragraph.Descendants<A.RunProperties>().Any(properties => properties.FontSize?.Value == 2800));
        Assert.Contains(generatedParagraphs, paragraph => paragraph.InnerText == "Editable text." && paragraph.Descendants<A.RunProperties>().Any(properties => properties.FontSize?.Value == 1800));
        var parentBullet = Assert.Single(generatedParagraphs, paragraph => paragraph.InnerText.Contains("Parent", StringComparison.Ordinal));
        var childBullet = Assert.Single(generatedParagraphs, paragraph => paragraph.InnerText.Contains("Child", StringComparison.Ordinal));
        Assert.Equal(0, parentBullet.ParagraphProperties?.Level?.Value);
        Assert.Equal(1, childBullet.ParagraphProperties?.Level?.Value);
        Assert.Contains(parentBullet.Descendants<A.RunProperties>(), properties => properties.Bold?.Value == true);
        Assert.Contains(childBullet.Descendants<A.RunProperties>(), properties => properties.Italic?.Value == true);
        var picture = Assert.Single(slides[0].Slide.Descendants<P.Picture>());
        var pictureExtents = picture.ShapeProperties!.Transform2D!.Extents!;
        Assert.Equal(pictureExtents.Cx!.Value, pictureExtents.Cy!.Value);
        Assert.InRange(pictureExtents.Cx.Value, 0L, 4000000L);
        Assert.InRange(pictureExtents.Cy.Value, 0L, 4500000L);
        var pictureOffset = picture.ShapeProperties.Transform2D.Offset!;
        Assert.InRange(pictureOffset.X!.Value, 5500000L, 9500000L - pictureExtents.Cx.Value);
        Assert.InRange(pictureOffset.Y!.Value, 1000000L, 5500000L - pictureExtents.Cy.Value);
        var generatedTable = Assert.Single(slides[1].Slide.Descendants<A.Table>());
        Assert.Equal(3, generatedTable.Elements<A.TableRow>().Count());
        Assert.Equal(new long[] { 3000000L, 3000000L }, generatedTable.GetFirstChild<A.TableGrid>()!.Elements<A.GridColumn>().Select(column => column.Width!.Value));
        Assert.All(generatedTable.Elements<A.TableRow>().First().Elements<A.TableCell>(), cell => Assert.Contains(cell.Descendants<A.RunProperties>(), properties => properties.Bold?.Value == true));
        Assert.All(generatedTable.Elements<A.TableRow>().Skip(1).SelectMany(row => row.Elements<A.TableCell>()), cell => Assert.Contains(cell.Descendants<A.RunProperties>(), properties => properties.Italic?.Value == true));
        Assert.All(slides, slide => Assert.DoesNotContain(slide.Slide.Descendants<A.Text>(), text => text.Text.Contains("{{", StringComparison.Ordinal)));
    }

    [Fact]
    public void Presentation_preflight_rejects_bad_markers_assets_and_tables_without_mutating_output()
    {
        using var fixture = PresentationFixture.CreateRich();
        var source = Path.Combine(fixture.Root, "bad.md");
        var output = Path.Combine(fixture.Root, "existing.pptx");
        var sentinel = new byte[] { 1, 2, 3, 4 };
        File.WriteAllBytes(output, sentinel);

        File.WriteAllText(source, "# Data {#table}\n\n| A | B | C |\n|---|---|---|\n| 1 | 2 | 3 |");
        var tableDiagnostics = PresentationAuthoring.Build(fixture.Template, new[] { new PresentationSource(source, File.ReadAllText(source)) }, fixture.Root, output);
        Assert.Contains(tableDiagnostics, diagnostic => diagnostic.Code == "YADG-PPT-TABLE-001");
        Assert.Equal(sentinel, File.ReadAllBytes(output));

        File.WriteAllText(source, "# Mixed {#text-image}\n\n## Picture {#image}\n\n![Missing](missing.png)");
        var assetDiagnostics = PresentationAuthoring.Build(fixture.Template, new[] { new PresentationSource(source, File.ReadAllText(source)) }, fixture.Root, output);
        Assert.Contains(assetDiagnostics, diagnostic => diagnostic.Code == "YADG-PPT-ASSET-001");
        Assert.Equal(sentinel, File.ReadAllBytes(output));

        using (var template = OpenXmlPresentationDocument.Open(fixture.Template, true))
        {
            var prototype = template.PresentationPart!.SlideParts.First(slide => slide.Slide.InnerText.Contains("{{slide:text-image}}", StringComparison.Ordinal));
            var marker = Assert.Single(prototype.Slide.Descendants<A.Text>(), text => text.Text == "{{slide:text-image}}");
            marker.Text = "{{slide:text image}}";
            prototype.Slide.Save();
        }
        File.WriteAllText(source, "# A slide\n\nContent.");
        var markerDiagnostics = PresentationAuthoring.Build(fixture.Template, new[] { new PresentationSource(source, File.ReadAllText(source)) }, fixture.Root, output);
        Assert.Contains(markerDiagnostics, diagnostic => diagnostic.Code == "YADG-PPT-TEMPLATE-003");
        Assert.Equal(sentinel, File.ReadAllBytes(output));
    }

    [Fact]
    public void Direct_presentation_commands_ignore_malformed_workspace_config_and_preserve_output_on_failure()
    {
        using var fixture = PresentationFixture.Create();
        var source = Path.Combine(fixture.Root, "slides.md");
        File.WriteAllText(source, "# Direct\n\nDirect body.");
        var secondSource = Path.Combine(fixture.Root, "slides-second.md");
        File.WriteAllText(secondSource, "# Direct second\n\nSecond body.");
        var output = Path.Combine(fixture.Root, "out", "direct.pptx");
        var staleDirectory = Path.Combine(fixture.Root, "YadgPresentations");
        Directory.CreateDirectory(staleDirectory);
        var stale = Path.Combine(staleDirectory, "keep.pptx");
        File.WriteAllText(stale, "configured deck stays isolated");

        var noConfigCheck = RunCli(fixture.Root, "presentation", "check", "--template", fixture.Template, "--source", source);
        Assert.Equal(0, noConfigCheck.ExitCode);
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "not valid front matter");
        var check = RunCli(fixture.Root, "presentation", "check", "--template", fixture.Template, "--source", source);
        Assert.Equal(0, check.ExitCode);
        Assert.Contains("presentation check: valid", check.StandardOutput, StringComparison.Ordinal);
        Assert.False(File.Exists(output));
        var orderedCheck = RunCli(fixture.Root, "presentation", "check", "--template", fixture.Template, "--source", source, "--source", secondSource);
        Assert.Equal(0, orderedCheck.ExitCode);
        Assert.Contains("valid (2 slide(s))", orderedCheck.StandardOutput, StringComparison.Ordinal);
        Assert.NotEqual(0, RunCli(fixture.Root, "presentation", "check", "--template", fixture.Template, "--source", source, "--unexpected").ExitCode);
        Assert.NotEqual(0, RunCli(fixture.Root, "presentation", "check", "--template", fixture.Template, "--template", fixture.Template, "--source", source).ExitCode);
        Assert.NotEqual(0, RunCli(fixture.Root, "presentation", "check", "--template", fixture.Template, "--source", Path.Combine(fixture.Root, "missing.md")).ExitCode);
        Assert.NotEqual(0, RunCli(fixture.Root, "presentation", "check", "--template", fixture.Template, "--source", source, "--output", output).ExitCode);

        var build = RunCli(fixture.Root, "presentation", "build", "--template", fixture.Template, "--source", source, "--source", secondSource, "--output", output);
        Assert.Equal(0, build.ExitCode);
        using (var builtPresentation = OpenXmlPresentationDocument.Open(output, false))
            Assert.Equal(2, builtPresentation.PresentationPart!.Presentation.SlideIdList!.Elements<P.SlideId>().Count());
        var originalOutput = File.ReadAllBytes(output);
        var originalStale = File.ReadAllBytes(stale);
        var duplicateOutput = RunCli(fixture.Root, "presentation", "build", "--template", fixture.Template, "--source", source, "--output", output, "--output", Path.Combine(fixture.Root, "unexpected.pptx"));
        Assert.NotEqual(0, duplicateOutput.ExitCode);
        Assert.Equal(originalOutput, File.ReadAllBytes(output));
        File.WriteAllText(source, "# Unsupported {#unknown-layout}\n\nBody.");

        var failedBuild = RunCli(fixture.Root, "presentation", "build", "--template", fixture.Template, "--source", source, "--output", output);

        Assert.NotEqual(0, failedBuild.ExitCode);
        Assert.Equal(originalOutput, File.ReadAllBytes(output));
        Assert.Equal(originalStale, File.ReadAllBytes(stale));
    }

    [Fact]
    public void Direct_document_commands_author_one_atomic_unfinalized_output_with_explicit_config()
    {
        using var workspace = WorkspaceFixture.Create();
        var template = Path.Combine(workspace.Root, "template.docx");
        CreateDocxTemplate(template, "{{content:document}}");
        var source = Path.Combine(workspace.Root, "content.md");
        File.WriteAllText(source, "# Document {#document}\n\nHello {{value:name}}.");
        var config = Path.Combine(workspace.Root, "authoring.yaml");
        File.WriteAllText(config, "yadg:\n  version: 1\nvalues:\n  name: \"Ada\"\n");
        var output = Path.Combine(workspace.Root, "out", "report.docx");
        File.WriteAllText(Path.Combine(workspace.Root, "YADG.md"), "malformed ambient workspace marker");

        var check = RunCli(workspace.Root, "document", "check", "--template", template, "--source", source, "--config", config);
        Assert.Equal(0, check.ExitCode);
        Assert.False(File.Exists(output));

        var build = RunCli(workspace.Root, "document", "build", "--template", template, "--source", source, "--output", output, "--config", config);
        Assert.True(build.ExitCode == 0, build.StandardOutput + Environment.NewLine + build.StandardError);
        Assert.Contains("unfinalized DOCX", build.StandardOutput, StringComparison.Ordinal);
        Assert.False(Directory.Exists(Path.Combine(workspace.Root, "YadgPreWords")));
        byte[] previous = File.ReadAllBytes(output);
        using (var authored = WordprocessingDocument.Open(output, false))
            Assert.Contains(authored.MainDocumentPart!.Document.Body!.Descendants<W.Text>(), text => text.Text.Contains("Hello Ada.", StringComparison.Ordinal));

        File.WriteAllText(source, "# Document {#document}\n\n![missing](missing.png){#figure}");
        var failed = RunCli(workspace.Root, "document", "build", "--template", template, "--source", source, "--output", output, "--config", config);
        Assert.NotEqual(0, failed.ExitCode);
        Assert.Equal(previous, File.ReadAllBytes(output));
    }

    [Fact]
    public void Selected_presentation_checks_and_builds_only_the_named_deck()
    {
        using var fixture = PresentationFixture.Create();
        var templates = Path.Combine(fixture.Root, "YadgTemplates");
        Directory.CreateDirectory(templates);
        File.Copy(fixture.Template, Path.Combine(templates, "main.pptx"));
        var source = Path.Combine(fixture.Root, "slides.md");
        File.WriteAllText(source, "# Selected\n\nSelected content.");
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\npresentations:\n  - id: selected\n    template: YadgTemplates/main.pptx\n    sources:\n      - slides.md\n    output: selected.pptx\n  - id: unrelated\n    template: YadgTemplates/missing.pptx\n    sources:\n      - missing.md\n    output: unrelated.pptx\n---\n");
        var outputDirectory = Path.Combine(fixture.Root, "YadgPresentations");
        Directory.CreateDirectory(outputDirectory);
        var selectedOutput = Path.Combine(outputDirectory, "selected.pptx");
        File.WriteAllText(selectedOutput, "old selected deck");
        var unrelatedOutput = Path.Combine(outputDirectory, "unrelated.pptx");
        File.WriteAllText(unrelatedOutput, "unrelated deck stays byte-identical");
        var unrelatedBefore = File.ReadAllBytes(unrelatedOutput);

        var check = RunCli(fixture.Root, "presentation", "check", "--id", "selected", "--workspace", fixture.Root);
        Assert.True(check.ExitCode == 0, check.StandardOutput + Environment.NewLine + check.StandardError);
        var build = RunCli(fixture.Root, "presentation", "build", "--id", "selected", "--workspace", fixture.Root);
        Assert.True(build.ExitCode == 0, build.StandardOutput + Environment.NewLine + build.StandardError);
        Assert.NotEqual("old selected deck", File.ReadAllText(selectedOutput));
        Assert.Equal(unrelatedBefore, File.ReadAllBytes(unrelatedOutput));
        Assert.NotEqual(0, RunCli(fixture.Root, "presentation", "check", "--id", "not-declared", "--workspace", fixture.Root).ExitCode);
    }

    [Fact]
    public void Pptx_yolo_missing_image_uses_a_visible_region_placeholder()
    {
        using var fixture = PresentationFixture.CreateRich();
        var source = Path.Combine(fixture.Root, "slides.md");
        File.WriteAllText(source, "# Diagram {#text-image}\n\n## Notes {#text}\n\nDiagram follows.\n\n## Picture {#image}\n\n![overview](missing.png)");
        var output = Path.Combine(fixture.Root, "out", "placeholder.pptx");

        var strict = RunCli(fixture.Root, "presentation", "build", "--template", fixture.Template, "--source", source, "--output", output);
        Assert.NotEqual(0, strict.ExitCode);
        Assert.False(File.Exists(output));

        var recovered = RunCli(fixture.Root, "presentation", "build", "--template", fixture.Template, "--source", source, "--output", output, "--yolo");
        Assert.True(recovered.ExitCode == 0, recovered.StandardOutput + Environment.NewLine + recovered.StandardError);
        Assert.Contains("degradations=1", recovered.StandardOutput, StringComparison.Ordinal);
        using var deck = OpenXmlPresentationDocument.Open(output, false);
        Assert.Contains(deck.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>(), shape => shape.TextBody?.InnerText.Contains("YADG image unavailable", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void Mermaid_image_region_uses_direct_configured_producer_and_builds_a_native_picture()
    {
        using var fixture = PresentationFixture.CreateRich();
        var source = Path.Combine(fixture.Root, "slides.md");
        File.WriteAllText(source, "# Diagram {#text-image}\n\n## Notes {#text}\n\nDiagram follows.\n\n## Picture {#image}\n\n```mermaid {#architecture caption=\"Architecture\"}\nflowchart LR\n  A --> B\n```");
        var script = Path.Combine(fixture.Root, "emit-png.ps1");
        File.WriteAllText(script, "param([string]$i,[string]$o)\n[IO.File]::WriteAllBytes($o, [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII='))\n");
        var config = Path.Combine(fixture.Root, "authoring.yaml");
        var powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "WindowsPowerShell", "v1.0", "powershell.exe").Replace('\\', '/');
        File.WriteAllText(config, $"yadg:\n  version: 1\nproducers:\n  mermaid:\n    executable: \"{powershell}\"\n    arguments:\n      - \"-NoProfile\"\n      - \"-File\"\n      - \"emit-png.ps1\"\n");
        var output = Path.Combine(fixture.Root, "out", "mermaid.pptx");

        var check = RunCli(fixture.Root, "presentation", "check", "--template", fixture.Template, "--source", source, "--config", config);
        Assert.True(check.ExitCode == 0, check.StandardOutput + Environment.NewLine + check.StandardError);
        Assert.False(File.Exists(output));
        var build = RunCli(fixture.Root, "presentation", "build", "--template", fixture.Template, "--source", source, "--output", output, "--config", config);
        Assert.True(build.ExitCode == 0, build.StandardOutput + Environment.NewLine + build.StandardError);
        using var deck = OpenXmlPresentationDocument.Open(output, false);
        var slide = deck.PresentationPart!.SlideParts.Single();
        Assert.Single(slide.Slide.Descendants<P.Picture>());
        Assert.DoesNotContain(slide.Slide.Descendants<P.Shape>(), shape => shape.TextBody?.InnerText.Contains("Architecture", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void Configured_build_failure_preserves_the_complete_previous_presentation_set()
    {
        using var workspace = WorkspaceFixture.Create();
        var templateDirectory = Path.Combine(workspace.Root, "YadgTemplates");
        Directory.CreateDirectory(templateDirectory);
        var template = Path.Combine(templateDirectory, "bad.pptx");
        File.WriteAllText(template, "not a PowerPoint package");
        var sourceDirectory = Path.Combine(workspace.Root, "slides");
        Directory.CreateDirectory(sourceDirectory);
        File.WriteAllText(Path.Combine(sourceDirectory, "deck.md"), "# Slide\n\nBody.");
        File.WriteAllText(Path.Combine(workspace.Root, "YADG.md"), "---\nyadg:\n  version: 1\npresentations:\n  - id: valid-id\n    template: YadgTemplates/bad.pptx\n    sources:\n      - slides/deck.md\n    output: new.pptx\n---\n");
        var outputDirectory = Path.Combine(workspace.Root, "YadgPresentations");
        Directory.CreateDirectory(outputDirectory);
        File.WriteAllText(Path.Combine(outputDirectory, "old.pptx"), "old one");
        File.WriteAllText(Path.Combine(outputDirectory, "keep.pptx"), "old two");
        var before = Directory.EnumerateFiles(outputDirectory, "*.pptx").ToDictionary(path => Path.GetFileName(path)!, File.ReadAllBytes, StringComparer.OrdinalIgnoreCase);

        var result = RunCli(workspace.Root, "build", "--workspace", workspace.Root);

        Assert.NotEqual(0, result.ExitCode);
        var after = Directory.EnumerateFiles(outputDirectory, "*.pptx").ToDictionary(path => Path.GetFileName(path)!, File.ReadAllBytes, StringComparer.OrdinalIgnoreCase);
        Assert.Equal(before.Keys.Order(StringComparer.OrdinalIgnoreCase), after.Keys.Order(StringComparer.OrdinalIgnoreCase));
        foreach (var name in before.Keys) Assert.Equal(before[name], after[name]);
    }

    [Fact]
    public void Configured_build_replaces_the_output_set_with_multiple_named_decks()
    {
        using var fixture = PresentationFixture.Create();
        var source = Path.Combine(fixture.Root, "slides.md");
        File.WriteAllText(source, "# Shared slide\n\nShared content.");
        var templateDirectory = Path.Combine(fixture.Root, "YadgTemplates");
        Directory.CreateDirectory(templateDirectory);
        File.Copy(fixture.Template, Path.Combine(templateDirectory, "shared.pptx"));
        var templateRelative = "YadgTemplates/shared.pptx";
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), $"---\nyadg:\n  version: 1\npresentations:\n  - id: overview\n    template: {templateRelative}\n    sources:\n      - slides.md\n    output: overview.pptx\n  - id: details\n    template: {templateRelative}\n    sources:\n      - slides.md\n    output: details.pptx\n---\n");
        var outputDirectory = Path.Combine(fixture.Root, "YadgPresentations");
        Directory.CreateDirectory(outputDirectory);
        File.WriteAllText(Path.Combine(outputDirectory, "obsolete.pptx"), "old result");

        var result = RunCli(fixture.Root, "build", "--workspace", fixture.Root);

        Assert.True(result.ExitCode == 0, result.StandardOutput + Environment.NewLine + result.StandardError);
        Assert.Equal(new[] { "details.pptx", "overview.pptx" }, Directory.EnumerateFiles(outputDirectory, "*.pptx").Select(Path.GetFileName).Order(StringComparer.OrdinalIgnoreCase));
        foreach (var name in new[] { "overview.pptx", "details.pptx" })
        {
            using var presentation = OpenXmlPresentationDocument.Open(Path.Combine(outputDirectory, name), false);
            Assert.Single(presentation.PresentationPart!.Presentation.SlideIdList!.Elements<P.SlideId>());
        }
    }

    [Fact]
    public void Configured_docx_and_pptx_result_sets_commit_independently()
    {
        using var workspace = WorkspaceFixture.Create();
        var templates = Path.Combine(workspace.Root, "YadgTemplates");
        Directory.CreateDirectory(templates);
        CreateDocxTemplate(Path.Combine(templates, "report.docx"), "{{content:document}}");
        File.WriteAllText(Path.Combine(templates, "broken.pptx"), "not a PowerPoint package");
        File.WriteAllText(Path.Combine(workspace.Root, "document.md"), "# Document {#document}\n\nFresh DOCX output.");
        File.WriteAllText(Path.Combine(workspace.Root, "slides.md"), "# Slide\n\nPPTX source excluded from DOCX discovery.");
        File.WriteAllText(Path.Combine(workspace.Root, "YADG.md"), "---\nyadg:\n  version: 1\npresentations:\n  - id: overview\n    template: YadgTemplates/broken.pptx\n    sources:\n      - slides.md\n    output: overview.pptx\n---\n");

        var docxOutput = Path.Combine(workspace.Root, "YadgPreWords", "report.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(docxOutput)!);
        File.WriteAllText(docxOutput, "old DOCX sentinel");
        var pptxOutput = Path.Combine(workspace.Root, "YadgPresentations");
        Directory.CreateDirectory(pptxOutput);
        File.WriteAllText(Path.Combine(pptxOutput, "old.pptx"), "old PPTX sentinel");

        var result = RunCli(workspace.Root, "build", "--workspace", workspace.Root);

        Assert.NotEqual(0, result.ExitCode);
        using (var authored = WordprocessingDocument.Open(docxOutput, false))
            Assert.Contains(authored.MainDocumentPart!.Document.Body!.Descendants<W.Text>(), text => text.Text == "Fresh DOCX output.");
        Assert.Single(Directory.EnumerateFiles(Path.GetDirectoryName(docxOutput)!, "*.docx", SearchOption.TopDirectoryOnly));
        Assert.False(File.Exists(Path.Combine(Path.GetDirectoryName(docxOutput)!, "slides.docx")));
        Assert.Equal("old PPTX sentinel", File.ReadAllText(Path.Combine(pptxOutput, "old.pptx")));
        Assert.Single(Directory.EnumerateFiles(pptxOutput, "*.pptx", SearchOption.TopDirectoryOnly));
    }

    private static void CreateDocxTemplate(string path, string contentTag)
    {
        using var document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var styles = main.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new W.Styles(new W.Style { Type = W.StyleValues.Paragraph, StyleId = "BodyText" });
        styles.Styles.Append(Enumerable.Range(1, 6).Select(index => new W.Style { Type = W.StyleValues.Paragraph, StyleId = $"Heading{index}" }));
        styles.Styles.Save();
        main.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Heading1" }, new W.OutlineLevel { Val = 0 }), new W.Run(new W.Text("Template heading"))),
            new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = "BodyText" }), new W.Run(new W.Text(contentTag))),
            new W.Paragraph(new W.Run(new W.Text("Template static paragraph")))));
        main.Document.Save();
    }

    private static (int ExitCode, string StandardOutput, string StandardError) RunCli(string workingDirectory, params string[] arguments)
    {
        var start = new ProcessStartInfo("dotnet") { WorkingDirectory = workingDirectory, UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
        start.ArgumentList.Add(typeof(Yadg.Cli.Program).Assembly.Location);
        foreach (var argument in arguments) start.ArgumentList.Add(argument);
        using var process = Process.Start(start)!;
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return (process.ExitCode, stdout, stderr);
    }

    private sealed class PresentationFixture : IDisposable
    {
        private PresentationFixture(string root, string template) { Root = root; Template = template; }
        public string Root { get; }
        public string Template { get; }

        public static PresentationFixture Create()
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0018-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            var template = Path.Combine(root, "template.pptx");
            using var document = OpenXmlPresentationDocument.Create(template, PresentationDocumentType.Presentation);
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new P.Presentation(new P.SlideMasterIdList(), new P.SlideIdList(), new P.SlideSize { Cx = 12192000, Cy = 6858000 });
            var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new P.SlideLayout(new P.CommonSlideData(CreateShapeTree()), new P.ColorMapOverride(new A.MasterColorMapping()));
            masterPart.SlideMaster = new P.SlideMaster(new P.CommonSlideData(CreateShapeTree()), new P.ColorMap(), new P.SlideLayoutIdList(), new P.TextStyles());
            masterPart.SlideMaster.SlideLayoutIdList!.Append(new P.SlideLayoutId { Id = 1U, RelationshipId = masterPart.GetIdOfPart(layoutPart) });
            presentationPart.Presentation.SlideMasterIdList!.Append(new P.SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(masterPart) });
            var slide = presentationPart.AddNewPart<SlidePart>();
            slide.AddPart(layoutPart);
            var tree = CreateShapeTree();
            tree.Append(Shape(2, "prototype", "{{slide:content}}", 0, 0, 1000000, 1000000));
            tree.Append(Shape(3, "title", "{{title}}", 0, 0, 9000000, 800000));
            tree.Append(TextShape(4, "content", 0, 1000000, 9000000, 4500000,
                new A.Paragraph(new A.Run(new A.Text("{{content}}"))),
                new A.Paragraph(new A.Run(new A.RunProperties { FontSize = 1800 }, new A.Text("paragraph example")))));
            tree.Append(Shape(5, "static", "static template element", 0, 6000000, 1000000, 300000));
            slide.Slide = new P.Slide(new P.CommonSlideData(tree), new P.ColorMapOverride(new A.MasterColorMapping()));
            var slideId = presentationPart.GetIdOfPart(slide);
            presentationPart.Presentation.SlideIdList!.Append(new P.SlideId { Id = 256U, RelationshipId = slideId });
            presentationPart.Presentation.Save(); masterPart.SlideMaster.Save(); layoutPart.SlideLayout.Save(); slide.Slide.Save();
            return new(root, template);
        }

        public static PresentationFixture CreateRich()
        {
            var fixture = Create();
            using var document = OpenXmlPresentationDocument.Open(fixture.Template, true);
            var presentationPart = document.PresentationPart!;
            var layoutPart = presentationPart.SlideMasterParts.Single().SlideLayoutParts.Single();
            AddPrototype(presentationPart, layoutPart, "text-image", new OpenXmlElement[]
            {
                Shape(2, "prototype", "{{slide:text-image}}", 0, 0, 1000000, 1000000),
                Shape(3, "title", "{{title}}", 0, 0, 9000000, 800000),
                TextShape(4, "text", 0, 1000000, 5000000, 4500000,
                    new A.Paragraph(new A.Run(new A.Text("{{text}}"))),
                    new A.Paragraph(new A.Run(new A.RunProperties { FontSize = 1800 }, new A.Text("paragraph example"))),
                    new A.Paragraph(new A.Run(new A.RunProperties { FontSize = 2800 }, new A.Text("heading example"))),
                    new A.Paragraph(new A.ParagraphProperties(new A.CharacterBullet { Char = "•" }), new A.Run(new A.RunProperties { FontSize = 1800 }, new A.Text("bullet example")))),
                Shape(5, "image", "{{image}}", 5500000, 1000000, 4000000, 4500000)
            });
            AddPrototype(presentationPart, layoutPart, "table", new OpenXmlElement[]
            {
                Shape(2, "prototype", "{{slide:table}}", 0, 0, 1000000, 1000000),
                Shape(3, "title", "{{title}}", 0, 0, 9000000, 800000),
                TableFrame()
            });
            presentationPart.Presentation.Save();
            return fixture;
        }

        private static void AddPrototype(PresentationPart presentationPart, SlideLayoutPart layoutPart, string layout, IEnumerable<OpenXmlElement> content)
        {
            var slide = presentationPart.AddNewPart<SlidePart>();
            slide.AddPart(layoutPart);
            var tree = CreateShapeTree();
            foreach (var shape in content) tree.Append(shape);
            slide.Slide = new P.Slide(new P.CommonSlideData(tree), new P.ColorMapOverride(new A.MasterColorMapping()));
            slide.Slide.Save();
            var ids = presentationPart.Presentation.SlideIdList!;
            var id = ids.Elements<P.SlideId>().Select(existing => existing.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max() + 1;
            ids.Append(new P.SlideId { Id = id, RelationshipId = presentationPart.GetIdOfPart(slide) });
        }

        private static P.GraphicFrame TableFrame()
        {
            A.TableCell Cell(string text) => new(
                new A.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.RunProperties { Bold = text == "{{table}}" || text == "header example", Italic = text == "body example" }, new A.Text(text)))),
                new A.TableCellProperties());
            A.TableRow Row(string first, string second) => new(Cell(first), Cell(second)) { Height = 200000 };
            var table = new A.Table(new A.TableProperties(),
                new A.TableGrid(new A.GridColumn { Width = 3000000 }, new A.GridColumn { Width = 3000000 }),
                Row("{{table}}", "header example"), Row("body example", "body example"));
            return new P.GraphicFrame(
                new P.NonVisualGraphicFrameProperties(new P.NonVisualDrawingProperties { Id = 6U, Name = "table" }, new P.NonVisualGraphicFrameDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
                new P.Transform(new A.Offset { X = 0L, Y = 1000000L }, new A.Extents { Cx = 9000000L, Cy = 4500000L }),
                new A.Graphic(new A.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }));
        }

        private static P.ShapeTree CreateShapeTree() => new(
            new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 1U, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L }, new A.ChildOffset { X = 0L, Y = 0L }, new A.ChildExtents { Cx = 0L, Cy = 0L })));

        private static P.Shape Shape(uint id, string name, string text, long x, long y, long cx, long cy) => TextShape(id, name, x, y, cx, cy, new A.Paragraph(new A.Run(new A.Text(text))));

        private static P.Shape TextShape(uint id, string name, long x, long y, long cx, long cy, params A.Paragraph[] paragraphs) => new(
            new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties { Id = id, Name = name }, new P.NonVisualShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy })),
            CreateTextBody(paragraphs));

        private static P.TextBody CreateTextBody(IEnumerable<A.Paragraph> paragraphs)
        {
            var body = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
            foreach (var paragraph in paragraphs) body.Append(paragraph);
            return body;
        }

        public void Dispose() { try { if (Directory.Exists(Root)) Directory.Delete(Root, true); } catch { } }
    }

    private sealed class WorkspaceFixture : IDisposable
    {
        private WorkspaceFixture(string root) => Root = root;
        public string Root { get; }
        public static WorkspaceFixture Create()
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0018-workspace-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            return new(root);
        }
        public void Dispose() { try { if (Directory.Exists(Root)) Directory.Delete(Root, true); } catch { } }
    }

    private static string FindRepositoryRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
            if (File.Exists(Path.Combine(directory.FullName, "Yadg.slnx"))) return directory.FullName;
        throw new DirectoryNotFoundException("Could not find the YADG repository root from the integration-test output directory.");
    }
}
