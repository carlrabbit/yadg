using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;
using Yadg.Word;

namespace Yadg.IntegrationTests;

public sealed class M0010FocusedTests
{
    [Fact]
    public void Section_rebases_relative_to_effective_template_outline_and_preserves_body_prototype()
    {
        using var workspace = FocusedWorkspace.Create("{{section:architecture}}", outlineLevel: 3);
        var model = Parse("# Architecture {#architecture}\n\nFirst body.\n\n### Deep {#deep}\n\nSecond body.");
        var output = Path.Combine(workspace.Root, "authored.docx");

        WordAuthoring.Author(workspace.Template, output, model);

        using var document = WordprocessingDocument.Open(output, false);
        var paragraphs = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToArray();
        Assert.Contains(paragraphs, p => Text(p) == "Architecture" && Style(p) == "Heading5");
        Assert.Contains(paragraphs, p => Text(p) == "Deep" && Style(p) == "Heading7");
        Assert.Contains(paragraphs, p => Text(p) == "First body." && Style(p) == "BodyText");
        Assert.Contains(paragraphs, p => Text(p) == "Second body." && Style(p) == "BodyText");
    }

    [Fact]
    public void Content_omits_root_and_rebases_descendant_using_source_delta()
    {
        using var workspace = FocusedWorkspace.Create("{{content:architecture}}", outlineLevel: 3);
        var model = Parse("# Architecture {#architecture}\n\nBody.\n\n### Deep {#deep}\n\nDetails.");
        var output = Path.Combine(workspace.Root, "authored.docx");

        WordAuthoring.Author(workspace.Template, output, model);

        using var document = WordprocessingDocument.Open(output, false);
        var paragraphs = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToArray();
        Assert.DoesNotContain(paragraphs, p => Text(p) == "Architecture");
        Assert.Contains(paragraphs, p => Text(p) == "Deep" && Style(p) == "Heading6");
        Assert.Contains(paragraphs, p => Text(p) == "Body." && Style(p) == "BodyText");
    }

    [Fact]
    public void Effective_heading_overflow_fails_before_output_mutation()
    {
        using var workspace = FocusedWorkspace.Create("{{section:architecture}}", outlineLevel: 3);
        var model = Parse("# Architecture {#architecture}\n\n###### Too deep.");
        var output = Path.Combine(workspace.Root, "authored.docx");
        File.WriteAllText(output, "sentinel");

        var exception = Assert.Throws<WordAuthoringException>(() => WordAuthoring.Author(workspace.Template, output, model));

        Assert.Contains("YADG-WORD-HEADING-OVERFLOW", exception.Message);
        Assert.Equal("sentinel", File.ReadAllText(output));
    }

    [Fact]
    public void Application_fixture_values_are_replaced_across_supported_story_parts()
    {
        var fixture = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "fixtures", "m0010", "word-origin.docx");
        fixture = Path.GetFullPath(fixture);
        if (!File.Exists(fixture)) return;
        var root = Path.Combine(Path.GetTempPath(), "yadg-m0010-values-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var output = Path.Combine(root, "authored.docx");
            var model = Parse("# Architecture {#architecture}\n\nBody.");
            WordAuthoring.Author(fixture, output, model, new Dictionary<string, string> { ["story-value"] = "Focused value" });
            using var archive = ZipFile.OpenRead(output);
            var xml = string.Join("\n", archive.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).Select(e => { using var reader = new StreamReader(e.Open(), Encoding.UTF8); return reader.ReadToEnd(); }));
            Assert.DoesNotContain("{{value:", xml, StringComparison.Ordinal);
            Assert.Contains("Focused value", xml, StringComparison.Ordinal);
        }
        finally { Directory.Delete(root, true); }
    }

    private static YadgDocument Parse(string markdown)
    {
        var result = MarkdownDocumentParser.Parse(markdown, "m0010-focused.md");
        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Diagnostics));
        return result.Document!;
    }

    private static string Text(Paragraph paragraph) => string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
    private static string? Style(Paragraph paragraph) => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

    private sealed class FocusedWorkspace : IDisposable
    {
        public string Root { get; }
        public string Template { get; }
        private FocusedWorkspace(string root, string template) { Root = root; Template = template; }
        public static FocusedWorkspace Create(string tag, int outlineLevel)
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0010-focused-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            var template = Path.Combine(root, "template.docx");
            using var document = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document);
            var main = document.AddMainDocumentPart();
            var styles = new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "BodyText" });
            styles.Append(Enumerable.Range(1, 9).Select(i => new Style { Type = StyleValues.Paragraph, StyleId = $"Heading{i}" }));
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            main.StyleDefinitionsPart!.Styles!.Save();
            var body = new Body(
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading4" }, new OutlineLevel { Val = outlineLevel }), new Run(new Text("Template context"))),
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "BodyText" }), new Run(new Text(tag))),
                new Paragraph(new Run(new Text("Static after."))));
            main.Document = new Document(body);
            main.Document.Save();
            return new(root, template);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
}
