using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;
using Yadg.Cli;
using Yadg.Word;

namespace Yadg.IntegrationTests;

public sealed class M0014FocusedTests
{
    private const string NestedBoxes = "┌────────────┐\n│ Outer      │\n│ ┌──────┐   │\n│ │Inner │   │\n│ └──────┘   │\n└────────────┘";
    private const string OccludedBoxes = "      ┌───────┐\n      │ upper │\n┌─────│       │─────┐\n│lower│       │     │\n│     └───────┘     │\n│                   │\n│                   │\n│                   │\n└───────────────────┘";
    private const string UncoveredGap = "┌─────         ─────┐\n│                   │\n│                   │\n└───────────────────┘";

    [Fact]
    public void Explicit_diagram_requires_exact_stable_id_form_and_produces_figure()
    {
        var parsed = MarkdownDocumentParser.Parse($"```diagram {{#flow caption=\"Flow\"}}\n{NestedBoxes}\n```", "content.md");
        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        var figure = Assert.IsType<YadgFigure>(Assert.Single(parsed.Document!.Blocks));
        Assert.Equal("flow", figure.Id);
        Assert.Equal("Flow", figure.AltText);
        Assert.NotNull(figure.GeneratedSvg);

        foreach (var info in new[] { "diagram", "diagram {#bad extra=\"x\"}", "Diagram {#flow}", "diagram {#flow caption=\"x\" extra=\"y\"}" })
        {
            var invalid = MarkdownDocumentParser.Parse($"```{info}\n{NestedBoxes}\n```", "content.md");
            Assert.Contains(invalid.Diagnostics, diagnostic => diagnostic.IsError);
        }
        var empty = MarkdownDocumentParser.Parse("```diagram {#empty}\n\n```", "content.md");
        Assert.Contains(empty.Diagnostics, diagnostic => diagnostic.Code == "YADG-DIAGRAM-002");
        var plain = MarkdownDocumentParser.Parse("```diagram {#plain}\nonly text\n```", "content.md");
        Assert.Contains(plain.Diagnostics, diagnostic => diagnostic.Code == "YADG-DIAGRAM-002");
    }

    [Fact]
    public void Grid_parser_recognizes_unicode_ascii_text_boxes_and_emits_deterministic_safe_svg()
    {
        var unicode = TextDiagramParser.Parse(NestedBoxes);
        Assert.True(unicode.Success, unicode.Error);
        Assert.Equal(2, unicode.Model!.Rectangles.Count);
        Assert.Contains(unicode.Model.Cells, cell => cell.Glyph == '┌' && cell.Connections == 6);
        Assert.Contains("Outer", unicode.Model.Svg, StringComparison.Ordinal);
        Assert.Contains("Inner", unicode.Model.Svg, StringComparison.Ordinal);
        Assert.Equal(unicode.Model.Svg, TextDiagramParser.Parse(NestedBoxes).Model!.Svg);
        var occluded = TextDiagramParser.Parse(OccludedBoxes);
        Assert.True(occluded.Success, occluded.Error);
        Assert.False(occluded.Model!.HasInvalidOcclusion);
        Assert.Contains(occluded.Model.Rectangles, rectangle => rectangle.Reconstructed && rectangle.CoveredBy?.Count == 1);
        Assert.Contains("occlusion-", occluded.Model.Svg, StringComparison.Ordinal);
        var uncovered = TextDiagramParser.Parse(UncoveredGap);
        Assert.True(uncovered.Success, uncovered.Error);
        Assert.True(uncovered.Model!.HasInvalidOcclusion);
        Assert.False(uncovered.Model.HasStrongGeometry);

        var ascii = TextDiagramParser.Parse("+----+\n| A  |\n+----+");
        Assert.True(ascii.Success, ascii.Error);
        Assert.Single(ascii.Model!.Rectangles);
        var xml = XDocument.Parse(unicode.Model.Svg);
        XNamespace svg = "http://www.w3.org/2000/svg";
        Assert.True(double.IsFinite(double.Parse(xml.Root!.Attribute("width")!.Value, System.Globalization.CultureInfo.InvariantCulture)));
        Assert.NotEmpty(xml.Descendants(svg + "rect"));
        Assert.NotEmpty(xml.Descendants(svg + "text"));
        Assert.DoesNotContain(xml.Descendants(), node => node.Name.LocalName is "script" or "foreignObject");
        Assert.DoesNotContain(xml.Descendants().Attributes().Where(attribute => attribute.Name.LocalName is "href" or "src"), attribute => !string.IsNullOrEmpty(attribute.Value));
    }

    [Fact]
    public void Required_junction_arrow_and_ascii_compatibility_primitives_are_connected()
    {
        var junctions = TextDiagramParser.Parse("┌─┬─┐\n│ │ │\n├─┼─┤\n│ │ │\n└─┴─┘");
        Assert.True(junctions.Success, junctions.Error);
        foreach (var glyph in "─│┌┐└┘├┤┬┴┼") Assert.Contains(junctions.Model!.Cells, cell => cell.Glyph == glyph);
        var unicodeArrow = TextDiagramParser.Parse("───→");
        Assert.True(unicodeArrow.Model!.HasStrongGeometry);
        Assert.Contains("<path", unicodeArrow.Model.Svg, StringComparison.Ordinal);
        var asciiArrow = TextDiagramParser.Parse("--->");
        Assert.True(asciiArrow.Model!.HasStrongGeometry);
        Assert.Contains("<path", asciiArrow.Model.Svg, StringComparison.Ordinal);
        var punctuation = TextDiagramParser.Parse("line-item label");
        Assert.False(punctuation.Model!.HasStrongGeometry);
        Assert.Contains("line-item", punctuation.Model.Svg, StringComparison.Ordinal);
    }

    [Fact]
    public void Autodetection_is_opt_in_scoped_and_conservative()
    {
        var config = WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\ndiagrams:\n  autoDetect: true\n---\n", "YADG.md", new List<Diagnostic>());
        Assert.True(config.DiagramsAutoDetect);
        foreach (var invalidConfig in new[] { "---\nyadg:\n  version: 1\ndiagrams:\n  autoDetect: maybe\n---", "---\nyadg:\n  version: 1\ndiagrams:\n  other: true\n---" })
        {
            var diagnostics = new List<Diagnostic>();
            WorkspaceValuesParser.Parse(invalidConfig, "YADG.md", diagnostics);
            Assert.Contains(diagnostics, diagnostic => diagnostic.IsError);
        }

        var unicode = MarkdownDocumentParser.Parse($"```\n{NestedBoxes}\n```", "content.md", diagramsAutoDetect: true);
        Assert.IsType<YadgDiagram>(Assert.Single(unicode.Document!.Blocks));
        Assert.Empty(unicode.Document.Figures);
        var identified = MarkdownDocumentParser.Parse($"```text {{#detected caption=\"Detected\"}}\n{NestedBoxes}\n```", "content.md", diagramsAutoDetect: true);
        Assert.True(identified.IsValid, string.Join(Environment.NewLine, identified.Diagnostics));
        Assert.Equal("detected", Assert.IsType<YadgFigure>(Assert.Single(identified.Document!.Blocks)).Id);

        foreach (var weak in new[] { "A -> B", "key: value", "ordinary prose" })
        {
            var result = MarkdownDocumentParser.Parse($"```text\n{weak}\n```", "content.md", diagramsAutoDetect: true);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "YADG-MD-UNSUPPORTED");
            Assert.Empty(result.Document!.Figures);
        }
        var language = MarkdownDocumentParser.Parse($"```csharp\n{NestedBoxes}\n```", "content.md", diagramsAutoDetect: true);
        Assert.Contains(language.Diagnostics, diagnostic => diagnostic.Code == "YADG-MD-UNSUPPORTED");
        var disabled = MarkdownDocumentParser.Parse($"```\n{NestedBoxes}\n```", "content.md");
        Assert.Contains(disabled.Diagnostics, diagnostic => diagnostic.Code == "YADG-MD-UNSUPPORTED");
        var explicitStillWorks = MarkdownDocumentParser.Parse($"```diagram {{#explicit}}\n{NestedBoxes}\n```", "content.md");
        Assert.True(explicitStillWorks.IsValid);
        var asciiCandidate = MarkdownDocumentParser.Parse("```\n+----+\n| A  |\n+----+\n```", "content.md", diagramsAutoDetect: true);
        Assert.IsType<YadgDiagram>(Assert.Single(asciiCandidate.Document!.Blocks));
        var unicodeArrowCandidate = MarkdownDocumentParser.Parse("```\n───→\n```", "content.md", diagramsAutoDetect: true);
        Assert.IsType<YadgDiagram>(Assert.Single(unicodeArrowCandidate.Document!.Blocks));
        var uncovered = MarkdownDocumentParser.Parse($"```\n{UncoveredGap}\n```", "content.md", diagramsAutoDetect: true);
        Assert.Contains(uncovered.Diagnostics, diagnostic => diagnostic.Code == "YADG-MD-UNSUPPORTED");
    }

    [Fact]
    public void Check_and_build_reject_invalid_explicit_diagrams_before_output_mutation()
    {
        using var fixture = DiagramWorkspace.Create("{{section:architecture}}");
        var templates = Path.Combine(fixture.Root, "YadgTemplates");
        Directory.CreateDirectory(templates);
        File.Copy(fixture.Template, Path.Combine(templates, "template.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Architecture {#architecture}\n\n```diagram {#broken}\nplain text\n```\n");
        var error = new StringWriter();
        var previous = Console.Error;
        try
        {
            Console.SetError(error);
            Assert.NotEqual(0, Program.Main(new[] { "check", "--workspace", fixture.Root }));
            Assert.NotEqual(0, Program.Main(new[] { "build", "--workspace", fixture.Root }));
        }
        finally { Console.SetError(previous); }
        Assert.Contains("YADG-DIAGRAM-002", error.ToString(), StringComparison.Ordinal);
        Assert.Contains("content.md:3:1", error.ToString(), StringComparison.Ordinal);
        Assert.False(Directory.Exists(Path.Combine(fixture.Root, "YadgPreWords")));
    }

    [Fact]
    public void Generated_svg_is_embedded_in_word_and_file_figures_remain_png_jpeg_only()
    {
        using var fixture = DiagramWorkspace.Create("{{figure:flow}}");
        var parsed = MarkdownDocumentParser.Parse($"```diagram {{#flow}}\n{NestedBoxes}\n```", "content.md");
        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        var output = Path.Combine(fixture.Root, "authored.docx");
        WordAuthoring.Author(fixture.Template, output, parsed.Document!);

        using var document = WordprocessingDocument.Open(output, false);
        Assert.Contains(document.MainDocumentPart!.ImageParts, part => part.ContentType == "image/svg+xml");
        var xml = document.MainDocumentPart.Document.OuterXml;
        Assert.Contains("svgBlip", xml, StringComparison.Ordinal);
        Assert.Contains("flow", xml, StringComparison.Ordinal);
        var svgPart = document.MainDocumentPart.ImageParts.Single(part => part.ContentType == "image/svg+xml");
        using var svgStream = svgPart.GetStream();
        using var reader = new StreamReader(svgStream);
        Assert.Contains("Outer", reader.ReadToEnd(), StringComparison.Ordinal);
        var sourceSvg = Path.Combine(fixture.Root, "source.svg");
        File.WriteAllText(sourceSvg, "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1\" height=\"1\"/>");
        var fileFigure = new YadgFigure("source", "", "source.svg", Path.Combine(fixture.Root, "content.md"));
        var imageDocument = new YadgDocument(new YadgBlock[] { fileFigure }, new Dictionary<string, YadgSection>(), new Dictionary<string, YadgTable>(), new Dictionary<string, YadgFigure> { ["source"] = fileFigure });
        var imageDiagnostics = new List<Diagnostic>();
        AssetValidation.AttachAndValidate(imageDocument, fixture.Root, imageDiagnostics);
        Assert.Contains(imageDiagnostics, diagnostic => diagnostic.Code == "YADG-FIGURE-006");
    }

    [Fact]
    public void Anonymous_diagram_remains_unregistered_and_authors_at_natural_position()
    {
        using var fixture = DiagramWorkspace.Create("{{content:architecture}}");
        var parsed = MarkdownDocumentParser.Parse($"# Architecture {{#architecture}}\n\nBefore.\n\n```text\n{NestedBoxes}\n```\n\nAfter.", "content.md", diagramsAutoDetect: true);
        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        Assert.Empty(parsed.Document!.Figures);
        Assert.IsType<YadgDiagram>(parsed.Document.FindSection("architecture")!.Body[1]);
        var output = Path.Combine(fixture.Root, "anonymous.docx");
        WordAuthoring.Author(fixture.Template, output, parsed.Document);
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Single(document.MainDocumentPart!.ImageParts);
        Assert.Contains("Before.", document.MainDocumentPart.Document.Body!.InnerText, StringComparison.Ordinal);
        Assert.Contains("After.", document.MainDocumentPart.Document.Body.InnerText, StringComparison.Ordinal);
    }

    [Fact]
    public void Identified_diagram_supports_natural_and_direct_placement_without_double_emission()
    {
        var source = $"# Architecture {{#architecture}}\n\nBefore.\n\n```diagram {{#flow caption=\"Flow\"}}\n{NestedBoxes}\n```";
        var natural = MarkdownDocumentParser.Parse(source, "content.md");
        Assert.True(natural.IsValid, string.Join(Environment.NewLine, natural.Diagnostics));
        using (var fixture = DiagramWorkspace.Create("{{section:architecture}}"))
        {
            var naturalOutput = Path.Combine(fixture.Root, "natural.docx");
            WordAuthoring.Author(fixture.Template, naturalOutput, natural.Document!);
            using var doc = WordprocessingDocument.Open(naturalOutput, false);
            Assert.Single(doc.MainDocumentPart!.ImageParts);
            Assert.Contains("Flow", doc.MainDocumentPart.Document.Body!.InnerText, StringComparison.Ordinal);
        }
        using (var fixture = DiagramWorkspace.Create("{{section:architecture}}", "{{figure:flow}}"))
        {
            var directOutput = Path.Combine(fixture.Root, "direct.docx");
            WordAuthoring.Author(fixture.Template, directOutput, natural.Document!);
            using var doc = WordprocessingDocument.Open(directOutput, false);
            Assert.Single(doc.MainDocumentPart!.ImageParts);
        }
    }

    private sealed class DiagramWorkspace : IDisposable
    {
        public string Root { get; }
        public string Template { get; }
        private DiagramWorkspace(string root, string template) { Root = root; Template = template; }
        public static DiagramWorkspace Create(params string[] tags)
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0014-focused-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            var template = Path.Combine(root, "template.docx");
            using var document = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document);
            var main = document.AddMainDocumentPart();
            var styles = new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "Caption" });
            styles.Append(Enumerable.Range(1, 9).Select(level => new Style { Type = StyleValues.Paragraph, StyleId = $"Heading{level}" }));
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            main.StyleDefinitionsPart!.Styles!.Save();
            main.Document = new Document(new Body(
                tags.Select(tag => (OpenXmlElement)new Paragraph(new Run(new Text(tag)))).Append(new SectionProperties(new PageSize { Width = 12240, Height = 15840 }, new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }))));
            main.Document.Save();
            return new(root, template);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
}
