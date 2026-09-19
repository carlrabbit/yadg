using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;

namespace Yadg.IntegrationTests;

public sealed class M0003IntegrationTests
{
    [Fact]
    public void Structured_workspace_builds_lists_table_png_jpeg_and_per_template_placements()
    {
        using var workspace = StructuredWorkspace.Create();
        var loaded = WorkspaceLoader.Load(workspace.Root);
        Assert.True(loaded.Document.Tables.ContainsKey("matrix"), string.Join(Environment.NewLine, loaded.Diagnostics) + " keys=" + string.Join(",", loaded.Document.Tables.Keys) + " blocks=" + string.Join(",", loaded.Document.Blocks.Select(b => b.GetType().Name)));
        Assert.Contains("context-figure", loaded.Document.Figures.Keys);
        var check = RunCli(workspace.Root, "check");
        Assert.True(check.ExitCode == 0, check.Output);
        var build = RunCli(workspace.Root, "build");
        Assert.True(build.ExitCode == 0, build.Output);

        var natural = Inspect(Path.Combine(workspace.Root, "YadgPreWords", "natural.docx"));
        Assert.Equal(1, natural.Tables);
        Assert.Equal(2, natural.Images);
        Assert.Contains(5943600L, natural.Extents);
        Assert.Contains("ListBullet", natural.Styles);
        Assert.Contains("ListNumber", natural.Styles);
        Assert.Contains("Caption", natural.Styles);
        Assert.Contains("TableGrid", natural.Styles);

        var relocated = Inspect(Path.Combine(workspace.Root, "YadgPreWords", "relocated.docx"));
        Assert.Equal(1, relocated.Tables);
        Assert.Equal(2, relocated.Images);
        Assert.Contains("Caption", relocated.Styles);
    }

    [Fact]
    public void Global_object_identity_and_invalid_structured_inputs_are_diagnosed()
    {
        using var workspace = StructuredWorkspace.Create();
        File.WriteAllText(Path.Combine(workspace.Root, "collision.md"), "# Collision {#matrix}\n\nText.");
        var duplicate = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, duplicate.ExitCode);
        Assert.Contains("YADG-REF-002", duplicate.Output);

        File.Delete(Path.Combine(workspace.Root, "collision.md"));
        File.WriteAllText(Path.Combine(workspace.Root, "unsupported.md"), "## Bad {#bad}\n\n- parent\n  - nested\n\n| A | B |\n|---|---|\n| 1 | 2 |");
        var invalid = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, invalid.ExitCode);
        Assert.Contains("YADG-LIST-001", invalid.Output);
        Assert.Contains("YADG-TABLE-001", invalid.Output);

        File.Delete(Path.Combine(workspace.Root, "unsupported.md"));
        CreateTemplate(Path.Combine(workspace.Root, "YadgTemplates", "mismatch.docx"), "{{content:content}}", ("{{table:context-figure}}", "{{figure:context-figure}}"));
        using (var duplicateDocument = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgTemplates", "mismatch.docx"), true))
        {
            duplicateDocument.MainDocumentPart!.Document.Body!.Append(new Paragraph(new Run(new Text("{{figure:context-figure}}"))));
            duplicateDocument.MainDocumentPart.Document.Save();
        }
        var placement = RunCli(workspace.Root, "check");
        Assert.Contains("YADG-REF-003", placement.Output);
        Assert.Contains("YADG-PLACEMENT-002", placement.Output);
    }

    [Fact]
    public void Figure_asset_paths_and_formats_are_validated_before_output_changes()
    {
        using var workspace = StructuredWorkspace.Create();
        File.WriteAllBytes(Path.Combine(workspace.Root, "images", "bad.gif"), new byte[] { 1, 2, 3 });
        File.WriteAllText(Path.Combine(workspace.Root, "bad.md"), "## Bad {#bad}\n\n![Bad](images/bad.gif){#bad-figure}");
        var output = Path.Combine(workspace.Root, "YadgPreWords", "natural.docx");
        File.WriteAllText(output, "sentinel");
        var result = RunCli(workspace.Root, "build");
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("YADG-FIGURE-006", result.Output);
        Assert.Equal("sentinel", File.ReadAllText(output));
    }

    private static (int Tables, int Images, string[] Styles, string FirstStructuredParagraph, string[] FigureNames, long[] Extents) Inspect(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var paragraphs = body.Descendants<Paragraph>().ToArray();
        var text = paragraphs.Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))).ToArray();
        var styles = paragraphs.SelectMany(p => p.Descendants<ParagraphStyleId>()).Select(s => s.Val?.Value ?? "").Concat(body.Descendants<TableStyle>().Select(s => s.Val?.Value ?? "")).ToArray();
        return (body.Descendants<Table>().Count(), document.MainDocumentPart.ImageParts.Count(), styles, text.FirstOrDefault(t => t.Contains("First", StringComparison.Ordinal)) ?? "", body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().Select(p => p.Name?.Value ?? "").ToArray(), body.Descendants().Where(e => e.LocalName == "extent").Select(e => long.TryParse(e.GetAttribute("cx", "").Value, out var value) ? value : 0L).ToArray());
    }

    private static ProcessResult RunCli(string root, string command)
    {
        var info = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "yadg.exe")) { WorkingDirectory = root, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        info.ArgumentList.Add(command); var process = Process.Start(info)!; process.WaitForExit(); return new(process.ExitCode, process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd());
    }

    private sealed record ProcessResult(int ExitCode, string Output);

    private sealed class StructuredWorkspace : IDisposable
    {
        public string Root { get; }
        private StructuredWorkspace(string root) => Root = root;
        public static StructuredWorkspace Create()
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0003-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "YadgTemplates")); Directory.CreateDirectory(Path.Combine(root, "YadgPreWords")); Directory.CreateDirectory(Path.Combine(root, "images"));
            File.WriteAllText(Path.Combine(root, "YADG.md"), "Synthetic structured workspace.");
            File.WriteAllText(Path.Combine(root, "content.md"), "## Content {#content}\n\n- First\n- Second with **strong**\n\n1. One\n2. Two\n\n| Name | Protocol |\n|---|---|\n| Alpha | HTTPS |\n| Beta | SFTP |\n{#matrix}\n\n![System context](images/context.png){#context-figure}\n\n![JPEG logo](images/logo.jpg){#logo-figure}");
            File.WriteAllBytes(Path.Combine(root, "images", "context.png"), LargePng());
            File.WriteAllBytes(Path.Combine(root, "images", "logo.jpg"), Jpeg());
            CreateTemplate(Path.Combine(root, "YadgTemplates", "natural.docx"), "{{content:content}}", null);
            CreateTemplate(Path.Combine(root, "YadgTemplates", "relocated.docx"), "{{content:content}}", ("{{table:matrix}}", "{{figure:context-figure}}"));
            return new(root);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }

    private static void CreateTemplate(string path, string contentTag, (string Table, string Figure)? direct)
    {
        using var document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "BodyText" }, new Style { Type = StyleValues.Paragraph, StyleId = "Caption" }, new Style { Type = StyleValues.Table, StyleId = "TableGrid" });
        styles.Append(Enumerable.Range(1, 6).Select(i => new Style { Type = StyleValues.Paragraph, StyleId = $"Heading{i}" }));
        styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "ListBullet", StyleParagraphProperties = new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 1 })) });
        styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "ListNumber", StyleParagraphProperties = new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 2 })) });
        stylesPart.Styles = styles; stylesPart.Styles.Save();
        var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(new AbstractNum { AbstractNumberId = 1 }, new AbstractNum { AbstractNumberId = 2 }, new NumberingInstance { NumberID = 1, AbstractNumId = new AbstractNumId { Val = 1 } }, new NumberingInstance { NumberID = 2, AbstractNumId = new AbstractNumId { Val = 2 } }); numberingPart.Numbering.Save();
        var body = new Body(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("Prepared template"))), new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "BodyText" }), new Run(new Text(contentTag))));
        if (direct is not null) { body.Append(new Paragraph(new Run(new Text(direct.Value.Table))), new Paragraph(new Run(new Text(direct.Value.Figure)))); }
        body.Append(new Paragraph(new Run(new Text("Unrelated template content"))), new SectionProperties(new PageSize { Width = 12240, Height = 15840 }, new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }));
        main.Document = new Document(body); main.Document.Save();
    }

    private static byte[] LargePng()
    {
        var bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        bytes[16] = 0x00; bytes[17] = 0x00; bytes[18] = 0x07; bytes[19] = 0xD0; bytes[20] = 0x00; bytes[21] = 0x00; bytes[22] = 0x03; bytes[23] = 0xE8; return bytes;
    }

    private static byte[] Jpeg() => Convert.FromBase64String("/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAH/AP/EABQQAQAAAAAAAAAAAAAAAAAAABD/2gAIAQEAAQUCcf/EABQRAQAAAAAAAAAAAAAAAAAAABD/2gAIAQMBAT8Bf//EABQRAQAAAAAAAAAAAAAAAAAAABD/2gAIAQIBAT8Bf//EABQQAQAAAAAAAAAAAAAAAAAAABD/2gAIAQEABj8Cf//Z");
}
