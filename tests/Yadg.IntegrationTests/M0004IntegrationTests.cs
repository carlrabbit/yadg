using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;

namespace Yadg.IntegrationTests;

public sealed class M0004IntegrationTests
{
    [Fact]
    public void Front_matter_prototypes_and_references_author_real_docx_structures()
    {
        using var workspace = M0004Workspace.Create();
        var loaded = WorkspaceLoader.Load(workspace.Root);
        Assert.True(loaded.IsValid, string.Join(Environment.NewLine, loaded.Diagnostics));
        Assert.Contains(loaded.Document.Blocks.OfType<YadgParagraph>().SelectMany(p => p.Inlines), inline => inline is YadgReference);
        var check = RunCli(workspace.Root, "check");
        Assert.True(check.ExitCode == 0, check.Output);
        var build = RunCli(workspace.Root, "build");
        Assert.True(build.ExitCode == 0, build.Output);

        using var output = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgPreWords", "references.docx"), false);
        var body = output.MainDocumentPart!.Document.Body!;
        var paragraphs = body.Descendants<Paragraph>().ToArray();
        var text = paragraphs.Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))).ToArray();
        Assert.DoesNotContain(text, value => value.Contains("yadg:frontmatter", StringComparison.Ordinal));
        Assert.DoesNotContain(text, value => value.Contains("yadg:prototype", StringComparison.Ordinal));
        Assert.Contains(text, value => value.Contains("Figure ", StringComparison.Ordinal) && value.Contains("Context", StringComparison.Ordinal));
        Assert.Contains(text, value => value.Contains("Table ", StringComparison.Ordinal) && value.Contains("Interfaces", StringComparison.Ordinal));
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("SEQ Figure", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("SEQ Table", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("REF yadg_fig", StringComparison.Ordinal) == true);
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("REF yadg_matrix", StringComparison.Ordinal) == true);
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("\\p", StringComparison.Ordinal) == true);
        Assert.NotEmpty(body.Descendants<BookmarkStart>());
        Assert.Equal(body.Descendants<BookmarkStart>().Count(), body.Descendants<BookmarkEnd>().Count());
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("TOC", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(paragraphs, p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "AltHeading");
    }

    [Fact]
    public void Invalid_front_matter_prototypes_and_numeric_targets_fail_before_output_changes()
    {
        using var workspace = M0004Workspace.Create();
        var output = Path.Combine(workspace.Root, "YadgPreWords", "references.docx"); File.WriteAllText(output, "sentinel");
        File.AppendAllText(Path.Combine(workspace.Root, "content.md"), "\n\nMissing [@missing].\n");
        var duplicate = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, duplicate.ExitCode);
        Assert.Contains("YADG-REF-004", duplicate.Output);
        Assert.Equal("sentinel", File.ReadAllText(output));

        File.WriteAllText(Path.Combine(workspace.Root, "YadgTemplates", "invalid.docx"), "not a docx");
        var invalid = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, invalid.ExitCode);
        Assert.Contains("YADG-WORD-OPEN", invalid.Output);
    }

    private static ProcessResult RunCli(string root, string command)
    {
        var info = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "yadg.exe")) { WorkingDirectory = root, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        info.ArgumentList.Add(command); using var process = Process.Start(info)!; process.WaitForExit(); return new(process.ExitCode, process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd());
    }

    private sealed record ProcessResult(int ExitCode, string Output);

    private sealed class M0004Workspace : IDisposable
    {
        public string Root { get; }
        private M0004Workspace(string root) => Root = root;
        public static M0004Workspace Create()
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0004-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "YadgTemplates")); Directory.CreateDirectory(Path.Combine(root, "YadgPreWords")); Directory.CreateDirectory(Path.Combine(root, "images"));
            File.WriteAllText(Path.Combine(root, "YADG.md"), "Synthetic M0004 workspace.");
            File.WriteAllText(Path.Combine(root, "content.md"), "## Architecture {#architecture}\n\nSee Section [@architecture], Figure [@fig], and Table [@matrix].\n\n![Context](images/context.png){#fig}\n\n| Name | Protocol |\n|---|---|\n| Alpha | HTTPS |\n{#matrix caption=\"Interfaces\"}");
            File.WriteAllBytes(Path.Combine(root, "images", "context.png"), Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
            CreateTemplate(Path.Combine(root, "YadgTemplates", "references.docx")); return new(root);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }

    private static void CreateTemplate(string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart(); var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "BodyText" }, new Style { Type = StyleValues.Paragraph, StyleId = "Caption" }, new Style { Type = StyleValues.Table, StyleId = "TableGrid" });
        styles.Append(Enumerable.Range(1, 6).Select(i => new Style { Type = StyleValues.Paragraph, StyleId = $"Heading{i}" })); styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "AltHeading", StyleParagraphProperties = new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 11 })) }); stylesPart.Styles = styles; stylesPart.Styles.Save();
        var numberingPart = main.AddNewPart<NumberingDefinitionsPart>(); numberingPart.Numbering = new Numbering(new AbstractNum { AbstractNumberId = 11 }, new NumberingInstance { NumberID = 11, AbstractNumId = new AbstractNumId { Val = 11 } }); numberingPart.Numbering.Save();
        var body = new Body();
        body.Append(P("{{yadg:frontmatter}}"), P("version: 1"), P("styles:"), P("  headings:"), P("    2: AltHeading"), P("  caption: Caption"), P("prototypes:"), P("  figureCaption: fig-prototype"), P("  tableCaption: table-prototype"), P("{{/yadg:frontmatter}}"));
        body.Append(P("{{yadg:prototype:fig-prototype}}"));
        body.Append(new Paragraph(
            new Run(new Text("Figure ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" SEQ Figure \\* ARABIC ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
            new Run(new Text(": ")), new Run(new Text("{{")), new Run(new Text("caption}}"))));
        body.Append(P("{{/yadg:prototype:fig-prototype}}"), P("{{yadg:prototype:table-prototype}}"));
        body.Append(new Paragraph(new Run(new Text("Table ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" SEQ Table \\* ARABIC ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }), new Run(new Text(": ")), new Run(new Text("{{caption}}"))));
        body.Append(P("{{/yadg:prototype:table-prototype}}"));
        body.Append(P("{{section:architecture}}"));
        body.Append(new Paragraph(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" TOC \\o \"1-3\" ")), new Run(new FieldChar { FieldCharType = FieldCharValues.End })));
        body.Append(new SectionProperties(new PageSize { Width = 12240, Height = 15840 }, new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }));
        main.Document = new Document(body); main.Document.Save();
    }

    private static Paragraph P(string text) => new(new Run(new Text(text)));

}
