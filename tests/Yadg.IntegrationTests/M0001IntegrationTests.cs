using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;
using Yadg.Word;

namespace Yadg.IntegrationTests;

public sealed class M0001IntegrationTests
{
    [Fact]
    public void Section_content_excludes_heading_and_keeps_nested_sections()
    {
        var parsed = MarkdownDocumentParser.Parse("## Architecture {#architecture}\n\nIntro.\n\n### Components\n\nDetails.\n\n## Next {#next}\n\nOther.");
        Assert.True(parsed.IsValid);
        var section = parsed.Document!.FindSection("architecture")!;
        Assert.Equal("Intro.\r\n\r\n### Components\r\n\r\nDetails.", section.ContentText);
        Assert.DoesNotContain("Architecture", section.ContentText);
    }

    [Fact]
    public void Split_run_tag_is_reconstructed_and_replaced_in_real_docx()
    {
        var root = Path.Combine(Path.GetTempPath(), "yadg-m0001-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var template = Path.Combine(root, "template.docx");
        var output = Path.Combine(root, "output.docx");
        CreateSyntheticTemplate(template);
        var parsed = MarkdownDocumentParser.Parse("## Architecture {#architecture}\n\nResolved content.");

        Assert.Equal(new[] { "{{yadg:section:content:architecture}}" }, WordAuthoring.FindTags(template));
        WordAuthoring.Author(template, output, parsed.Document!);

        using var document = WordprocessingDocument.Open(output, false);
        var paragraphs = document.MainDocumentPart!.Document.Body!.Descendants<Paragraph>().ToArray();
        var text = paragraphs.Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))).ToArray();
        Assert.Contains(text, value => value.Contains("Resolved content.", StringComparison.Ordinal));
        Assert.Contains("Unrelated template paragraph", text);
        Assert.DoesNotContain(text, value => value.Contains("{{yadg:", StringComparison.Ordinal));
        Assert.Equal("Heading1", paragraphs[0].ParagraphProperties!.ParagraphStyleId!.Val!.Value);
    }

    [Fact]
    public void Unresolved_reference_is_actionable()
    {
        var parsed = MarkdownDocumentParser.Parse("## Present {#present}\n\nText.");
        var result = MarkdownDocumentParser.Resolve(parsed.Document!, new SectionReference("missing", SectionSelection.Content));
        Assert.NotNull(result.Diagnostic);
        Assert.Equal("YADG-REF-001", result.Diagnostic!.Code);
        Assert.Contains("missing", result.Diagnostic.Message);
    }

    [Fact]
    public void Malformed_tag_is_actionable()
    {
        Assert.False(SectionReference.TryParseTag("{{yadg:section:content:}}", out _, out var diagnostic));
        Assert.Equal("YADG-TAG-001", diagnostic!.Code);
    }

    private static void CreateSyntheticTemplate(string path)
    {
        using var document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("Template heading"))),
            new Paragraph(
                new Run(new Text("Prefix ")),
                new Run(new RunProperties(new RunStyle { Val = "Emphasis" }), new Text("{{yadg:section:")),
                new Run(new Text("content:arch")),
                new Run(new Text("itecture}}")),
                new Run(new Text(" suffix"))),
            new Paragraph(new Run(new Text("Unrelated template paragraph")))));
        main.Document.Save();
    }
}
