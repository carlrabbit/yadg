using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Yadg.IntegrationTests;

public sealed class M0007IntegrationTests
{
    [Fact]
    public void Prepared_table_clones_rows_positionally_preserves_template_rows_and_authors_inlines()
    {
        using var workspace = TestWorkspace.Create("| Name | Detail |\n|---|---|\n| Alpha | *one* |\n| Beta | **two** |\n{#matrix caption=\"Interfaces\"}");
        var check = RunCli(workspace.Root, "check");
        Assert.Equal(0, check.ExitCode);
        var build = RunCli(workspace.Root, "build");
        Assert.Equal(0, build.ExitCode);

        using var output = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgPreWords", "prepared.docx"), false);
        var body = output.MainDocumentPart!.Document.Body!;
        var table = body.Elements<Table>().Single();
        var rows = table.Elements<TableRow>().ToArray();
        Assert.Equal(4, rows.Length); // template header + two data rows + template-owned following row
        Assert.Equal("Name", Logical(rows[0].Elements<TableCell>().First()));
        Assert.Equal("Detail", Logical(rows[0].Elements<TableCell>().Last()));
        Assert.Equal("[Alpha]", Logical(rows[1].Elements<TableCell>().First()));
        Assert.Equal("one", Logical(rows[1].Elements<TableCell>().Last()));
        Assert.Equal("[Beta]", Logical(rows[2].Elements<TableCell>().First()));
        Assert.Equal("two", Logical(rows[2].Elements<TableCell>().Last()));
        Assert.Equal("Following", Logical(rows[3]));
        Assert.DoesNotContain(body.Descendants<Text>(), t => t.Text?.Contains("{{", StringComparison.Ordinal) == true);
        Assert.Contains(body.Elements<Paragraph>(), p => Logical(p).Contains("Interfaces", StringComparison.Ordinal));
        Assert.Contains(rows[1].Descendants<Run>(), r => r.RunProperties?.Italic is not null);
        Assert.Contains(rows[2].Descendants<Run>(), r => r.RunProperties?.Bold is not null);
    }

    [Fact]
    public void Prepared_table_with_zero_body_rows_removes_controls_without_emitting_a_clone()
    {
        using var workspace = TestWorkspace.Create("| Name | Detail |\n|---|---|\n{#matrix}");
        Assert.Equal(0, RunCli(workspace.Root, "build").ExitCode);
        using var output = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgPreWords", "prepared.docx"), false);
        var rows = output.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToArray();
        Assert.Equal(2, rows.Length); // header and the following template-owned row
        Assert.Equal("Following", Logical(rows[1]));
        Assert.DoesNotContain(output.MainDocumentPart.Document.Body.Descendants<Text>(), t => t.Text?.Contains("{{", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void Prepared_table_conflict_and_malformed_prototype_fail_before_overwriting_output()
    {
        using var workspace = TestWorkspace.Create("| Name | Detail |\n|---|---|\n| Alpha | one |\n{#matrix}", includeGeneratedPlacement: true, malformedPrototype: true);
        var output = Path.Combine(workspace.Root, "YadgPreWords", "prepared.docx");
        File.WriteAllText(output, "sentinel");
        var check = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, check.ExitCode);
        Assert.Contains("YADG-PREPARED", check.Output);
        Assert.Contains("YADG-PLACEMENT-003", check.Output);
        Assert.Equal("sentinel", File.ReadAllText(output));
        var build = RunCli(workspace.Root, "build");
        Assert.NotEqual(0, build.ExitCode);
        Assert.Equal("sentinel", File.ReadAllText(output));
    }

    [Fact]
    public void Duplicate_prepared_bindings_fail_validation()
    {
        using var workspace = TestWorkspace.Create("| Name | Detail |\n|---|---|\n| Alpha | one |\n{#matrix}", duplicateBinding: true);
        var result = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("YADG-PLACEMENT-002", result.Output);
    }

    [Theory]
    [InlineData("tag")]
    [InlineData("nested")]
    [InlineData("grid")]
    [InlineData("vertical")]
    [InlineData("paragraph")]
    [InlineData("field")]
    [InlineData("bookmark")]
    [InlineData("drawing")]
    [InlineData("control")]
    public void Forbidden_prototype_structures_fail_validation(string variant)
    {
        using var workspace = TestWorkspace.Create("| Name | Detail |\n|---|---|\n| Alpha | one |\n{#matrix}", malformedPrototype: false, prototypeVariant: variant);
        var result = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("YADG-PREPARED", result.Output);
    }

    private static string Logical(OpenXmlElement element) => string.Concat(element.Descendants<Text>().Select(t => t.Text));
    private static ProcessResult RunCli(string root, string command)
    {
        var info = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "yadg.exe")) { WorkingDirectory = root, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        info.ArgumentList.Add(command); using var process = Process.Start(info)!; process.WaitForExit(); return new(process.ExitCode, process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd());
    }
    private sealed record ProcessResult(int ExitCode, string Output);

    private sealed class TestWorkspace : IDisposable
    {
        public string Root { get; }
        private TestWorkspace(string root) => Root = root;
        public static TestWorkspace Create(string markdown, bool includeGeneratedPlacement = false, bool malformedPrototype = false, bool duplicateBinding = false, string? prototypeVariant = null)
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0007-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "YadgTemplates")); Directory.CreateDirectory(Path.Combine(root, "YadgPreWords"));
            File.WriteAllText(Path.Combine(root, "YADG.md"), "Notes");
            File.WriteAllText(Path.Combine(root, "content.md"), markdown);
            CreateTemplate(Path.Combine(root, "YadgTemplates", "prepared.docx"), includeGeneratedPlacement, malformedPrototype, duplicateBinding, prototypeVariant);
            return new(root);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }

    private static void CreateTemplate(string path, bool includeGeneratedPlacement, bool malformedPrototype, bool duplicateBinding, string? prototypeVariant)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var styles = new Styles(new Style { Type = StyleValues.Table, StyleId = "TableGrid" }, new Style { Type = StyleValues.Paragraph, StyleId = "Caption" });
        main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
        var header = new TableRow(new TableCell(new Paragraph(new Run(new Text("Name")))), new TableCell(new Paragraph(new Run(new Text("Detail")))));
        var marker = new TableRow(new TableCell(new Paragraph(new Run(new Text("{{table-rows:matrix}}")))));
        var prototypeCell1 = new TableCell(new Paragraph(new Run(new RunProperties(new Color { Val = "FF0000" }), new Text("[")), new Run(new RunProperties(new Italic()), new Text("{{ce")), new Run(new RunProperties(new Italic()), new Text("ll}}")), new Run(new Text("]"))));
        var prototypeCell2 = malformedPrototype
            ? new TableCell(new Paragraph(new Run(new Text("{{cell}} {{cell}}"))))
            : prototypeVariant switch
            {
                "tag" => new TableCell(new Paragraph(new Run(new Text("{{cell}} {{value:missing}}")))),
                "nested" => new TableCell(new Paragraph(new Run(new Text("{{cell}}"))), new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("nested"))))))),
                "grid" => new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("{{cell}}")))),
                "vertical" => new TableCell(new TableCellProperties(new VerticalMerge()), new Paragraph(new Run(new Text("{{cell}}")))),
                "paragraph" => new TableCell(new Paragraph(new Run(new Text("{{cell}}"))), new Paragraph(new Run(new Text("extra")))),
                "field" => new TableCell(new Paragraph(new Run(new Text("{{cell}}")), new SimpleField { Instruction = " PAGE " })),
                "bookmark" => new TableCell(new Paragraph(new BookmarkStart { Id = "1", Name = "old" }, new Run(new Text("{{cell}}")), new BookmarkEnd { Id = "1" })),
                "drawing" => new TableCell(new Paragraph(new Run(new Text("{{cell}}"), new Drawing()))),
                "control" => new TableCell(new Paragraph(new Run(new Text("{{cell}}"))), new SdtBlock(new SdtContentBlock(new Paragraph(new Run(new Text("control")))))),
                _ => new TableCell(new Paragraph(new Run(new RunProperties(new Bold()), new Text("{{cell}}"))))
            };
        var prototype = new TableRow(prototypeCell1, prototypeCell2);
        var following = new TableRow(new TableCell(new Paragraph(new Run(new Text("Following"))), new TableCell(new Paragraph(new Run(new Text("row"))))));
        var prepared = new Table(new TableProperties(new TableStyle { Val = "TableGrid" }), new TableGrid(new GridColumn(), new GridColumn()), header, marker, prototype, following);
        var body = new Body(prepared);
        if (duplicateBinding)
        {
            body.Append(new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("{{table-rows:matrix}}"))))),
                new TableRow(new TableCell(new Paragraph(new Run(new Text("{{cell}}")))), new TableCell(new Paragraph(new Run(new Text("{{cell}}")))))));
        }
        if (includeGeneratedPlacement) body.Append(new Paragraph(new Run(new Text("{{table:matrix}}"))));
        body.Append(new SectionProperties());
        main.Document = new Document(body);
        main.Document.Save();
    }
}
