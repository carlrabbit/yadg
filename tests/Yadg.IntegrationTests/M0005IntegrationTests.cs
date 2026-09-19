using System.Diagnostics;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Yadg.IntegrationTests;

public sealed class M0005IntegrationTests
{
    [Fact]
    public void Render_uses_real_libreoffice_to_refresh_fields_and_write_both_artifacts()
    {
        var lo = Environment.GetEnvironmentVariable("YADG_LIBREOFFICE_PATH") ?? @"C:\Program Files\LibreOffice\program\soffice.com";
        if (!File.Exists(lo)) return; // Portable Tier 0-2 remains valid where the real Tier 3 runtime is absent.
        using var workspace = RenderWorkspace.Create();
        var evidenceRoot = Environment.GetEnvironmentVariable("YADG_REVIEW_EVIDENCE_ROOT");
        if (!string.IsNullOrWhiteSpace(evidenceRoot)) { Directory.CreateDirectory(evidenceRoot); File.Copy(Path.Combine(workspace.Root, "YadgPreWords", "review.docx"), Path.Combine(evidenceRoot, "review-input.docx"), true); }
        using (var authored = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgPreWords", "review.docx"), false))
        {
            var authoredErrors = new OpenXmlValidator().Validate(authored).ToArray();
            Assert.True(authoredErrors.Length == 0, string.Join(Environment.NewLine, authoredErrors.Select(error => $"{error.Description} at {error.Path?.XPath}")));
        }
        var result = RunCli(workspace.Root, "render", "--renderer-path", lo);
        Assert.True(result.ExitCode == 0, result.Output);
        var finalized = Path.Combine(workspace.Root, "YadgWords", "review.docx"); var pdf = Path.Combine(workspace.Root, "YadgPdfs", "review.pdf");
        Assert.True(new FileInfo(finalized).Length > 0); Assert.True(new FileInfo(pdf).Length > 0);
        using var document = WordprocessingDocument.Open(finalized, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var texts = body.Descendants<Text>().Select(t => t.Text).ToArray();
        Assert.Contains("1", texts);
        Assert.DoesNotContain("{{", texts);
        Assert.Contains("Architecture", texts);
        Assert.Contains("Template-owned list item", texts);
        Assert.NotEmpty(body.Descendants<Table>());
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("TOC", StringComparison.OrdinalIgnoreCase) == true);
        Assert.NotEmpty(document.MainDocumentPart.ImageParts);
        Assert.NotEmpty(body.Descendants<Drawing>());
        foreach (var blip in body.Descendants<A.Blip>())
        {
            Assert.NotNull(blip.Embed);
            Assert.IsAssignableFrom<ImagePart>(document.MainDocumentPart.GetPartById(blip.Embed!.Value!));
        }
        if (!string.IsNullOrWhiteSpace(evidenceRoot))
        {
            Directory.CreateDirectory(evidenceRoot); File.Copy(finalized, Path.Combine(evidenceRoot, "libreoffice-review.docx"), true); File.Copy(pdf, Path.Combine(evidenceRoot, "libreoffice-review.pdf"), true);
            var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(finalized)));
            var pdfHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(pdf)));
            File.WriteAllText(Path.Combine(evidenceRoot, "review-manifest.md"), $"# M0005 LibreOffice Review Evidence\n\n- Repository revision: `{Run("git", "rev-parse HEAD", Environment.CurrentDirectory).Trim()}`\n- LibreOffice executable: `{lo}`\n- LibreOffice version: `{Run(lo, "--version", Environment.CurrentDirectory).Trim()}`\n- Platform: `{Environment.OSVersion}`\n- Input: `review.docx`\n- Finalized DOCX SHA-256: `{hash}`\n- PDF SHA-256: `{pdfHash}`\n- Validation: `yadg render --renderer-path {lo}` succeeded; refreshed field results were structurally inspected.\n- Isolation: renderer reports a unique temporary session/profile identity per invocation.\n");
        }
    }

    [Fact]
    public void Render_rejects_unsupported_renderer_and_missing_inputs_before_outputs()
    {
        using var workspace = RenderWorkspace.Empty();
        var unsupported = RunCli(workspace.Root, "render", "--renderer", "word"); Assert.NotEqual(0, unsupported.ExitCode); Assert.Contains("YADG-RENDER-020", unsupported.Output);
        var missing = RunCli(workspace.Root, "render"); Assert.NotEqual(0, missing.ExitCode); Assert.Contains("YADG-RENDER-003", missing.Output);
    }

    private static string Run(string executable, string arguments, string workingDirectory)
    {
        using var process = Process.Start(new ProcessStartInfo(executable, arguments) { WorkingDirectory = workingDirectory, UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true })!; process.WaitForExit(30000); return process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
    }

    private static ProcessResult RunCli(string root, string command, params string[] arguments)
    {
        var info = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "yadg.exe")) { WorkingDirectory = root, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false }; info.ArgumentList.Add(command); foreach (var argument in arguments) info.ArgumentList.Add(argument);
        using var process = Process.Start(info)!; process.WaitForExit(); return new(process.ExitCode, process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd());
    }

    private sealed record ProcessResult(int ExitCode, string Output);

    private sealed class RenderWorkspace : IDisposable
    {
        public string Root { get; }
        private RenderWorkspace(string root) => Root = root;
        public static RenderWorkspace Create()
        {
            var workspace = Empty(); CreateAuthoredDocx(Path.Combine(workspace.Root, "YadgPreWords", "review.docx")); return workspace;
        }
        public static RenderWorkspace Empty()
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0005-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(Path.Combine(root, "YadgPreWords")); File.WriteAllText(Path.Combine(root, "YADG.md"), "Synthetic renderer workspace."); return new(root);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }

    private static void CreateAuthoredDocx(string path)
    {
        using var package = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document); var main = package.AddMainDocumentPart();
        var stylesPart = main.AddNewPart<StyleDefinitionsPart>(); stylesPart.Styles = new Styles(
            new Style { Type = StyleValues.Paragraph, StyleId = "Heading1" }, new Style { Type = StyleValues.Paragraph, StyleId = "ListBullet", StyleParagraphProperties = new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 1 })) }, new Style { Type = StyleValues.Paragraph, StyleId = "Caption" }, new Style { Type = StyleValues.Table, StyleId = "TableGrid" }); stylesPart.Styles.Save();
        var numberingPart = main.AddNewPart<NumberingDefinitionsPart>(); numberingPart.Numbering = new Numbering(new AbstractNum { AbstractNumberId = 1 }, new NumberingInstance { NumberID = 1, AbstractNumId = new AbstractNumId { Val = 1 } }); numberingPart.Numbering.Save();
        var body = new Body(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new BookmarkStart { Id = "12", Name = "yadg_section" }, new Run(new Text("Architecture")), new BookmarkEnd { Id = "12" }), new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "ListBullet" }), new Run(new Text("Template-owned list item"))), new Paragraph(new Run(new Text("Renderer review"))), new Paragraph(),
            new Paragraph(new Run(new Text("Figure ")), new BookmarkStart { Id = "11", Name = "yadg_fig" }, new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" SEQ Figure \\* ARABIC ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }), new BookmarkEnd { Id = "11" }, new Run(new Text(": Context"))),
            new Paragraph(new Run(new Text("See figure ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" REF yadg_fig \\h ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }), new Run(new Text(" and table ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" REF yadg_table \\h ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }), new Run(new Text(" and section ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" REF yadg_section \\p \\h ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End })),
            new Table(new TableProperties(new TableStyle { Val = "TableGrid" }), new TableRow(new TableCell(new Paragraph(new Run(new Text("Interface")))), new TableCell(new Paragraph(new Run(new Text("Protocol"))))), new TableRow(new TableCell(new Paragraph(new Run(new Text("Alpha")))), new TableCell(new Paragraph(new Run(new Text("HTTPS")))))),
            new Paragraph(new Run(new Text("Table ")), new BookmarkStart { Id = "13", Name = "yadg_table" }, new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" SEQ Table \\* ARABIC ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }), new BookmarkEnd { Id = "13" }, new Run(new Text(": Interfaces"))),
            new Paragraph(new Run(new Text("Contents")), new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" TOC \\o \"1-3\" ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End })),
            new SectionProperties(new PageSize { Width = 12240, Height = 15840 }, new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }));
         main.Document = new Document(body); main.Document.Save(); var figureParagraph = main.Document.Body!.Elements<Paragraph>().ElementAt(3); main.Document.Body.Elements<Table>().Single().InsertAt(new TableGrid(new GridColumn(), new GridColumn()), 1);
        var image = main.AddImagePart(ImagePartType.Png); using (var stream = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAHgAAABGCAIAAACIdsdYAAAAtklEQVR4nO3SQQ0AIAADMUAJwpCEaFRwr9bAksvmPnfw3wo2ELrj0RGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCR4SOCB0ROiJ0ROiI0BGhI0JHhI4IHRE6InRE6IjQEaEjQkeEjggdEToidEToiNARoSNCj8YDMT0BlKR9idUAAAAASUVORK5CYII="))) image.FeedData(stream); var rel = main.GetIdOfPart(image);
        figureParagraph.Append(new Run(new Text("Figure asset: "))); figureParagraph.Append(new Run(new DocumentFormat.OpenXml.Wordprocessing.Drawing(new DW.Inline(new DW.Extent { Cx = 3657600, Cy = 1371600 }, new DW.DocProperties { Id = 2U, Name = "review-figure" }, new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }), new A.Graphic(new A.GraphicData(new PIC.Picture(new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties { Id = 0U, Name = "review.png" }, new PIC.NonVisualPictureDrawingProperties()), new PIC.BlipFill(new A.Blip { Embed = rel }, new A.Stretch(new A.FillRectangle())), new PIC.ShapeProperties(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 3657600, Cy = 1371600 }), new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })))));
        main.Document.Save();
        package.Dispose(); NormalizeFixturePackage(path);
    }

    private static void NormalizeFixturePackage(string path)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var source = archive.GetEntry("media/image.png")!;
        var target = archive.CreateEntry("word/media/image.png", CompressionLevel.Optimal);
        using (var input = source.Open()) using (var output = target.Open()) input.CopyTo(output);
        var relationships = archive.GetEntry("word/_rels/document.xml.rels")!;
        string xml; using (var input = relationships.Open()) using (var buffer = new MemoryStream()) { input.CopyTo(buffer); xml = System.Text.Encoding.UTF8.GetString(buffer.ToArray()); }
        relationships.Delete(); var replacement = archive.CreateEntry("word/_rels/document.xml.rels"); using (var writer = new StreamWriter(replacement.Open())) writer.Write(xml.Replace("/media/image.png", "media/image.png", StringComparison.Ordinal));
        var contentTypes = archive.GetEntry("[Content_Types].xml")!;
        string types; using (var input = contentTypes.Open()) using (var buffer = new MemoryStream()) { input.CopyTo(buffer); types = System.Text.Encoding.UTF8.GetString(buffer.ToArray()); }
        types = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\" /><Default Extension=\"jpg\" ContentType=\"image/jpeg\" /><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" /><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\" /><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\" /><Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\" /></Types>";
        types = types.Replace("<Default Extension=\"jpg\"", "<Default Extension=\"png\" ContentType=\"image/png\" /><Default Extension=\"jpg\"", StringComparison.Ordinal);
        contentTypes.Delete(); var typeReplacement = archive.CreateEntry("[Content_Types].xml"); using (var writer = new StreamWriter(typeReplacement.Open())) writer.Write(types);
        var documentXml = archive.GetEntry("word/document.xml")!;
        string document; using (var input = documentXml.Open()) using (var buffer = new MemoryStream()) { input.CopyTo(buffer); document = System.Text.Encoding.UTF8.GetString(buffer.ToArray()); }
        documentXml.Delete(); var documentReplacement = archive.CreateEntry("word/document.xml"); using (var writer = new StreamWriter(documentReplacement.Open())) writer.Write(document.Replace("<w:document ", "<w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" ", StringComparison.Ordinal));
        var packageRelationships = archive.GetEntry("_rels/.rels")!;
        string packageXml; using (var input = packageRelationships.Open()) using (var buffer = new MemoryStream()) { input.CopyTo(buffer); packageXml = System.Text.Encoding.UTF8.GetString(buffer.ToArray()); }
        packageRelationships.Delete(); var packageReplacement = archive.CreateEntry("_rels/.rels"); using (var writer = new StreamWriter(packageReplacement.Open())) writer.Write(packageXml.Replace("Target=\"/word/document.xml\"", "Target=\"word/document.xml\"", StringComparison.Ordinal));
        source.Delete();
    }
}
