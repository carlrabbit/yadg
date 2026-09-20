using Yadg.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WordRendererEngine = Yadg.WordRenderer.WordRenderer;

namespace Yadg.IntegrationTests;

public sealed class M0009IntegrationTests
{
    [Fact]
    public void Publish_uses_configured_relative_destination_and_preserves_unrelated_files()
    {
        using var fixture = new PublishFixture();
        File.WriteAllText(Path.Combine(fixture.Root, "Published", "alpha.docx"), "old");
        var result = Publisher.Publish(fixture.Workspace);
        Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics));
        Assert.Equal("alpha", File.ReadAllText(Path.Combine(fixture.Root, "Published", "alpha.docx")));
        Assert.Equal("beta", File.ReadAllText(Path.Combine(fixture.Root, "Published", "beta.docx")));
        Assert.Equal("keep", File.ReadAllText(Path.Combine(fixture.Root, "Published", "keep.txt")));
        Assert.False(File.Exists(Path.Combine(fixture.Root, "Published", "nested.docx")));
        Assert.False(Directory.EnumerateFiles(Path.Combine(fixture.Root, "Published"), ".yadg-publish-*.tmp").Any());
    }

    [Fact]
    public void Publish_cli_destination_overrides_config_and_reserved_paths_fail_before_copy()
    {
        using var fixture = new PublishFixture();
        var overridePath = Path.Combine(fixture.Root, "Delivery");
        var result = Publisher.Publish(fixture.Workspace, overridePath);
        Assert.True(result.Success);
        Assert.True(File.Exists(Path.Combine(overridePath, "alpha.docx")));
        var reserved = Publisher.Publish(fixture.Workspace, Path.Combine(fixture.Root, "YadgWords", "subdir"));
        Assert.False(reserved.Success);
        Assert.Contains(reserved.Diagnostics, d => d.Code == "YADG-PUBLISH-002");
        Assert.False(Directory.Exists(Path.Combine(fixture.Root, "YadgWords", "subdir")));
    }

    [Fact]
    public void Publish_path_schema_accepts_path_and_rejects_unknown_keys()
    {
        var diagnostics = new List<Diagnostic>();
        var values = WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\npublish:\n  path: ./Published\n---", "YADG.md", diagnostics);
        Assert.DoesNotContain(diagnostics, d => d.IsError);
        Assert.Equal("./Published", values.PublishPath);
        diagnostics.Clear();
        WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\npublish:\n  other: ./Published\n---", "YADG.md", diagnostics);
        Assert.Contains(diagnostics, d => d.Code == "YADG-VALUES-007");
    }

    [Fact]
    public void Real_word_renderer_finalizes_a_docx_and_preserves_authored_input()
    {
        using var fixture = new WordFixture();
        var sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(fixture.Input)));
        var ownedBefore = System.Diagnostics.Process.GetProcessesByName("WINWORD").Select(p => p.Id).ToHashSet();
        var result = new WordRendererEngine().Render(fixture.Root);
        Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics));
        Assert.False(string.IsNullOrWhiteSpace(result.RuntimeVersion));
        Assert.Equal(sourceHash, Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(fixture.Input))));
        for (var attempt = 0; attempt < 20 && System.Diagnostics.Process.GetProcessesByName("WINWORD").Any(p => !ownedBefore.Contains(p.Id)); attempt++) Thread.Sleep(250);
        Assert.DoesNotContain(System.Diagnostics.Process.GetProcessesByName("WINWORD"), p => !ownedBefore.Contains(p.Id));
        using var finalized = WordprocessingDocument.Open(Path.Combine(fixture.Root, "YadgWords", "sample.docx"), false);
        Assert.Contains(finalized.MainDocumentPart!.Document.Body!.Descendants<Text>(), t => t.Text == "Word renderer fixture");
        Assert.Contains(finalized.MainDocumentPart.Document.Body.Descendants<FieldCode>(), f => f.Text?.Contains("SEQ Figure", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(finalized.MainDocumentPart.Document.Body.Descendants<FieldCode>(), f => f.Text?.Contains("REF figure-1", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(finalized.MainDocumentPart.Document.Body.Descendants<FieldCode>(), f => f.Text?.Contains("REF section-1", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(finalized.MainDocumentPart.Document.Body.Descendants<FieldCode>(), f => f.Text?.Contains("TOC", StringComparison.OrdinalIgnoreCase) == true);
        Assert.NotEmpty(finalized.MainDocumentPart.Document.Body.Descendants<Table>());
        Assert.NotEmpty(finalized.MainDocumentPart.ImageParts);
        var evidenceRoot = Environment.GetEnvironmentVariable("YADG_M0009_EVIDENCE_ROOT");
        if (!string.IsNullOrWhiteSpace(evidenceRoot))
        {
            Directory.CreateDirectory(evidenceRoot);
            File.Copy(fixture.Input, Path.Combine(evidenceRoot, "authored.docx"), true);
            File.Copy(Path.Combine(fixture.Root, "YadgWords", "sample.docx"), Path.Combine(evidenceRoot, "finalized.docx"), true);
            File.WriteAllText(Path.Combine(evidenceRoot, "word-version.txt"), result.RuntimeVersion);
        }
    }

    [Fact]
    public void Publish_command_uses_cli_override_without_building_or_rendering()
    {
        using var fixture = new CliPublishFixture();
        var info = new System.Diagnostics.ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "yadg.exe")) { WorkingDirectory = fixture.Root, UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
        info.ArgumentList.Add("publish"); info.ArgumentList.Add("--publish-path"); info.ArgumentList.Add(fixture.Override);
        using var process = System.Diagnostics.Process.Start(info)!;
        process.WaitForExit(30_000);
        var output = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
        Assert.Equal(0, process.ExitCode);
        Assert.Contains("copied 1 finalized DOCX", output);
        Assert.True(File.Exists(Path.Combine(fixture.Override, "final.docx")));
        Assert.False(Directory.Exists(Path.Combine(fixture.Root, "Configured")));
    }

    [Fact]
    public void Publish_supports_external_absolute_destination_and_rejects_missing_inputs_or_destination()
    {
        using var fixture = new PublishFixture();
        var external = Path.Combine(Path.GetTempPath(), "yadg-m0009-external-" + Guid.NewGuid().ToString("N"));
        try
        {
            var externalResult = Publisher.Publish(fixture.Workspace, external);
            Assert.True(externalResult.Success);
            Assert.Equal("alpha", File.ReadAllText(Path.Combine(external, "alpha.docx")));
            var noDestination = Publisher.Publish(fixture.Workspace with { Values = WorkspaceValues.Empty });
            Assert.False(noDestination.Success);
            Assert.Contains(noDestination.Diagnostics, d => d.Code == "YADG-PUBLISH-001");
            using var empty = new EmptyPublishFixture();
            var noInputs = Publisher.Publish(empty.Workspace, Path.Combine(empty.Root, "Delivery"));
            Assert.False(noInputs.Success);
            Assert.Contains(noInputs.Diagnostics, d => d.Code == "YADG-PUBLISH-004");
            Assert.False(Directory.Exists(Path.Combine(empty.Root, "Delivery")));
        }
        finally { try { Directory.Delete(external, true); } catch { } }
    }

    private sealed class PublishFixture : IDisposable
    {
        public string Root { get; } = Path.Combine(Path.GetTempPath(), "yadg-m0009-publish-" + Guid.NewGuid().ToString("N"));
        public YadgWorkspace Workspace { get; }

        public PublishFixture()
        {
            Directory.CreateDirectory(Path.Combine(Root, "YadgWords", "nested"));
            Directory.CreateDirectory(Path.Combine(Root, "Published"));
            File.WriteAllText(Path.Combine(Root, "YadgWords", "alpha.docx"), "alpha");
            File.WriteAllText(Path.Combine(Root, "YadgWords", "beta.docx"), "beta");
            File.WriteAllText(Path.Combine(Root, "YadgWords", "nested", "nested.docx"), "nested");
            File.WriteAllText(Path.Combine(Root, "Published", "keep.txt"), "keep");
            Workspace = new(Root, Array.Empty<string>(), Array.Empty<string>(), new YadgDocument(Array.Empty<YadgBlock>(), new Dictionary<string, YadgSection>(), new Dictionary<string, YadgTable>(), new Dictionary<string, YadgFigure>()), new WorkspaceValues(new Dictionary<string, string>()) { PublishPath = "./Published" }, Array.Empty<Diagnostic>());
        }

        public void Dispose() { try { Directory.Delete(Root, true); } catch { } }
    }

    private sealed class WordFixture : IDisposable
    {
        public string Root { get; } = Path.Combine(Path.GetTempPath(), "yadg-m0009-word-" + Guid.NewGuid().ToString("N"));
        public string Input => Path.Combine(Root, "YadgPreWords", "sample.docx");

        public WordFixture()
        {
            Directory.CreateDirectory(Path.Combine(Root, "YadgPreWords"));
            using var package = WordprocessingDocument.Create(Input, WordprocessingDocumentType.Document);
            var main = package.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "Heading1" }, new Style { Type = StyleValues.Paragraph, StyleId = "ListBullet" }, new Style { Type = StyleValues.Paragraph, StyleId = "Caption" }, new Style { Type = StyleValues.Table, StyleId = "TableGrid" });
            styles.Styles.Save();
            var numbering = main.AddNewPart<NumberingDefinitionsPart>();
            numbering.Numbering = new Numbering(new AbstractNum(new Level { LevelIndex = 0, NumberingFormat = new NumberingFormat { Val = NumberFormatValues.Decimal }, LevelText = new LevelText { Val = "%1." } }) { AbstractNumberId = 1 }, new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            numbering.Numbering.Save();
            var heading = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }, new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })), new BookmarkStart { Id = "1", Name = "section-1" }, new Run(new Text("Word renderer fixture")), new BookmarkEnd { Id = "1" });
            var figureCaption = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }), new Run(new Text("Figure ")), new BookmarkStart { Id = "2", Name = "figure-1" }, new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" SEQ Figure \\* ARABIC ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }), new BookmarkEnd { Id = "2" }, new Run(new Text(": Context figure")));
            var references = new Paragraph(new Run(new Text("See figure ")), Field(" REF figure-1 \\h "), new Run(new Text(" and section ")), Field(" REF section-1 \\p \\h "));
            var table = new Table(new TableProperties(new TableStyle { Val = "TableGrid" }), new TableRow(new TableCell(new Paragraph(new Run(new Text("Interface")))), new TableCell(new Paragraph(new Run(new Text("Protocol"))))), new TableRow(new TableCell(new Paragraph(new Run(new Text("Alpha")))), new TableCell(new Paragraph(new Run(new Text("HTTPS"))))));
            var tableCaption = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }), new Run(new Text("Table ")), new BookmarkStart { Id = "3", Name = "table-1" }, new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" SEQ Table \\* ARABIC ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }), new BookmarkEnd { Id = "3" }, new Run(new Text(": Interfaces")));
            var toc = new Paragraph(new Run(new Text("Contents")), Field(" TOC \\o \"1-3\" \\h "));
            var lof = new Paragraph(new Run(new Text("Figures")), Field(" TOC \\c \"Figure\" "));
            var lot = new Paragraph(new Run(new Text("Tables")), Field(" TOC \\c \"Table\" "));
            var figure = new Paragraph(new Run(new Text("Figure asset: ")));
            var body = new Body(heading, new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "ListBullet" }), new Run(new Text("Template-owned list item"))), figureCaption, references, table, tableCaption, toc, lof, lot, figure, new SectionProperties(new PageSize { Width = 12240, Height = 15840 }, new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }));
            main.Document = new Document(body);
            var image = main.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAHgAAABGCAIAAACIdsdYAAAAtklEQVR4nO3SQQ0AIAADMUAJwpCEaFRwr9bAksvmPnfw3wo2ELrj0RGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCj8YDMT0BlKR9idUAAAAASUVORK5CYII="))) image.FeedData(stream);
            var relation = main.GetIdOfPart(image);
            figure.Append(new Run(CreateFigure(relation)));
            main.Document.Save();
        }

        private static Run Field(string instruction) => new(new FieldChar { FieldCharType = FieldCharValues.Begin }, new FieldCode(instruction), new FieldChar { FieldCharType = FieldCharValues.Separate }, new Text("0"), new FieldChar { FieldCharType = FieldCharValues.End });

        private static DocumentFormat.OpenXml.Wordprocessing.Drawing CreateFigure(string relation) => new(new DW.Inline(
            new DW.Extent { Cx = 3657600, Cy = 1371600 },
            new DW.DocProperties { Id = 2U, Name = "review-figure" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(new PIC.Picture(
                new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties { Id = 0U, Name = "review.png" }, new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(new A.Blip { Embed = relation }, new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 3657600, Cy = 1371600 }), new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

        public void Dispose() { try { Directory.Delete(Root, true); } catch { } }
    }

    private sealed class CliPublishFixture : IDisposable
    {
        public string Root { get; } = Path.Combine(Path.GetTempPath(), "yadg-m0009-cli-publish-" + Guid.NewGuid().ToString("N"));
        public string Override => Path.Combine(Root, "Override");

        public CliPublishFixture()
        {
            Directory.CreateDirectory(Path.Combine(Root, "YadgTemplates"));
            Directory.CreateDirectory(Path.Combine(Root, "YadgWords"));
            File.WriteAllText(Path.Combine(Root, "YADG.md"), "---\nyadg:\n  version: 1\npublish:\n  path: ./Configured\n---\n");
            File.WriteAllText(Path.Combine(Root, "content.md"), "# Content {#content}\n");
            using (var package = WordprocessingDocument.Create(Path.Combine(Root, "YadgTemplates", "template.docx"), WordprocessingDocumentType.Document))
            {
                var main = package.AddMainDocumentPart(); main.Document = new Document(new Body(new Paragraph(new Run(new Text("template"))))); main.Document.Save();
            }
            File.Copy(Path.Combine(Root, "YadgTemplates", "template.docx"), Path.Combine(Root, "YadgWords", "final.docx"));
        }

        public void Dispose() { try { Directory.Delete(Root, true); } catch { } }
    }

    private sealed class EmptyPublishFixture : IDisposable
    {
        public string Root { get; } = Path.Combine(Path.GetTempPath(), "yadg-m0009-empty-publish-" + Guid.NewGuid().ToString("N"));
        public YadgWorkspace Workspace { get; }

        public EmptyPublishFixture()
        {
            Directory.CreateDirectory(Path.Combine(Root, "YadgWords"));
            Workspace = new(Root, Array.Empty<string>(), Array.Empty<string>(), new YadgDocument(Array.Empty<YadgBlock>(), new Dictionary<string, YadgSection>(), new Dictionary<string, YadgTable>(), new Dictionary<string, YadgFigure>()), WorkspaceValues.Empty, Array.Empty<Diagnostic>());
        }

        public void Dispose() { try { Directory.Delete(Root, true); } catch { } }
    }
}
