using System.Diagnostics;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Yadg.IntegrationTests;

public sealed class M0008IntegrationTests
{
    private const string Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

    [Fact]
    public void Mermaid_producer_renders_generated_figure_through_existing_word_path()
    {
        using var workspace = TestWorkspace.Create(configured: true);
        var check = RunCli(workspace.Root, "check");
        Assert.Equal(0, check.ExitCode);
        Assert.Contains("check: valid", check.Output);
        Assert.Empty(Directory.EnumerateFiles(workspace.Root, "*.png", SearchOption.AllDirectories));
        Assert.Equal(0, RunCli(workspace.Root, "build").ExitCode);
        using var output = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgPreWords", "mermaid.docx"), false);
        Assert.NotEmpty(output.MainDocumentPart!.ImageParts);
        var text = string.Concat(output.MainDocumentPart.Document.Body!.Descendants<Text>().Select(t => t.Text));
        Assert.DoesNotContain("flowchart", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("System flow", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Mermaid_without_configuration_fails_and_preserves_existing_output()
    {
        using var workspace = TestWorkspace.Create(configured: false);
        var output = Path.Combine(workspace.Root, "YadgPreWords", "mermaid.docx");
        File.WriteAllText(output, "sentinel");
        var result = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("YADG-PRODUCER-007", result.Output);
        Assert.Equal("sentinel", File.ReadAllText(output));
    }

    [Fact]
    public void Mermaid_input_output_arguments_are_reserved()
    {
        var diagnostics = new List<Yadg.Core.Diagnostic>();
        Yadg.Core.WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\nproducers:\n  mermaid:\n    executable: mmdc\n    arguments:\n      - --input\n---", "YADG.md", diagnostics);
        Assert.Contains(diagnostics, d => d.Code == "YADG-PRODUCER-006");
    }

    [Fact]
    public void Mermaid_process_start_exit_and_invalid_output_fail_actionably()
    {
        var root = Path.Combine(Path.GetTempPath(), "yadg-m0008-runner-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var figure = new Yadg.Core.YadgFigure("diagram", "", "generated.png", Path.Combine(root, "content.md"), null, "flowchart LR\nA-->B");
            var document = new Yadg.Core.YadgDocument(new Yadg.Core.YadgBlock[] { figure }, new Dictionary<string, Yadg.Core.YadgSection>(), new Dictionary<string, Yadg.Core.YadgTable>(), new Dictionary<string, Yadg.Core.YadgFigure> { [figure.Id] = figure });
            var values = new Yadg.Core.WorkspaceValues(new Dictionary<string, string>()) { Producers = new Dictionary<string, Yadg.Core.MermaidProducerConfiguration> { ["mermaid"] = new("missing-producer-executable", Array.Empty<string>()) } };
            var diagnostics = new List<Yadg.Core.Diagnostic>();
            Yadg.Core.MermaidProducer.RenderFigures(document, root, values, diagnostics, new List<string>());
            Assert.Contains(diagnostics, d => d.Code == "YADG-PRODUCER-010");

            values = values with { Producers = new Dictionary<string, Yadg.Core.MermaidProducerConfiguration> { ["mermaid"] = new("cmd.exe", new[] { "/d", "/c", "exit 7" }) } };
            diagnostics.Clear();
            Yadg.Core.MermaidProducer.RenderFigures(document, root, values, diagnostics, new List<string>());
            Assert.Contains(diagnostics, d => d.Code == "YADG-PRODUCER-012");

            File.WriteAllText(Path.Combine(root, "invalid.ps1"), "Set-Content -LiteralPath $args[3] -Value bad");
            values = values with { Producers = new Dictionary<string, Yadg.Core.MermaidProducerConfiguration> { ["mermaid"] = new("powershell.exe", new[] { "-NoProfile", "-File", "invalid.ps1" }) } };
            diagnostics.Clear();
            var temp = new List<string>();
            Yadg.Core.MermaidProducer.RenderFigures(document, root, values, diagnostics, temp);
            Assert.Contains(diagnostics, d => d.Code == "YADG-PRODUCER-013");
            foreach (var path in temp) try { Directory.Delete(path, true); } catch { }
        }
        finally { try { Directory.Delete(root, true); } catch { } }
    }

    [Fact]
    public void Tier3_real_bun_mermaid_path_authors_png_and_docx()
    {
        var bun = Environment.GetEnvironmentVariable("YADG_M0008_BUN");
        if (string.IsNullOrWhiteSpace(bun)) return;
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(FindRepositoryRoot(), "eng", "test-tools.json")));
        var bunVersion = manifest.RootElement.GetProperty("bun").GetProperty("version").GetString()!;
        var package = manifest.RootElement.GetProperty("mermaidCli").GetProperty("package").GetString()!;
        var packageVersion = manifest.RootElement.GetProperty("mermaidCli").GetProperty("version").GetString()!;
        using var workspace = TestWorkspace.CreateTier3(bun, package, packageVersion);
        Assert.Equal(0, RunCli(workspace.Root, "check").ExitCode);
        Assert.Equal(0, RunCli(workspace.Root, "build").ExitCode);
        using var output = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgPreWords", "mermaid.docx"), false);
        Assert.NotEmpty(output.MainDocumentPart!.ImageParts);
        var body = output.MainDocumentPart.Document.Body!;
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("SEQ Figure", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(body.Descendants<FieldCode>(), field => field.Text?.Contains("REF yadg_system_flow", StringComparison.OrdinalIgnoreCase) == true);
        Assert.NotEmpty(body.Descendants<BookmarkStart>());
        Assert.DoesNotContain(body.Descendants<Text>(), text => text.Text?.Contains("flowchart", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(bunVersion, RunProcess(bun, "--version").Output, StringComparison.Ordinal);
    }

    private static ProcessResult RunCli(string root, string command)
    {
        var info = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "yadg.exe")) { WorkingDirectory = root, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        info.ArgumentList.Add(command); info.ArgumentList.Add("--workspace"); info.ArgumentList.Add(root);
        using var process = Process.Start(info)!; process.WaitForExit(); return new(process.ExitCode, process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd());
    }

    private sealed record ProcessResult(int ExitCode, string Output);

    private sealed class TestWorkspace : IDisposable
    {
        public string Root { get; }
        private TestWorkspace(string root) => Root = root;
        public static TestWorkspace Create(bool configured)
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0008-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "YadgTemplates")); Directory.CreateDirectory(Path.Combine(root, "YadgPreWords"));
            var config = configured ? "producers:\n  mermaid:\n    executable: cmd.exe\n    arguments:\n      - /d\n      - /c\n      - powershell.exe\n      - -NoProfile\n      - -File\n      - fake-mermaid.ps1\n" : string.Empty;
            File.WriteAllText(Path.Combine(root, "YADG.md"), "---\nyadg:\n  version: 1\n" + config + "---\nNotes");
            if (configured) File.WriteAllText(Path.Combine(root, "fake-mermaid.ps1"), "[IO.File]::WriteAllBytes($args[3], [Convert]::FromBase64String('" + Png + "'))");
            File.WriteAllText(Path.Combine(root, "content.md"), "# Architecture {#architecture}\n\nIntro. See Figure [@system-flow].\n\n```mermaid {#system-flow caption=\"System flow\"}\nflowchart LR\n    A --> B\n```\n");
            CreateTemplate(Path.Combine(root, "YadgTemplates", "mermaid.docx"), includeNumericFigure: true);
            return new(root);
        }
        public static TestWorkspace CreateTier3(string bun, string package, string version)
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0008-tier3-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "YadgTemplates")); Directory.CreateDirectory(Path.Combine(root, "YadgPreWords"));
            var escapedBun = bun.Replace("\\", "\\\\", StringComparison.Ordinal);
            var config = "producers:\n  mermaid:\n    executable: \"" + escapedBun + "\"\n    arguments:\n      - x\n      - --bun\n      - --package\n      - \"" + package + "@" + version + "\"\n      - mmdc\n";
            File.WriteAllText(Path.Combine(root, "YADG.md"), "---\nyadg:\n  version: 1\n" + config + "---\nNotes");
            File.WriteAllText(Path.Combine(root, "content.md"), "# Architecture {#architecture}\n\nIntro. See Figure [@system-flow].\n\n```mermaid {#system-flow caption=\"System flow\"}\nflowchart LR\n    A --> B\n```\n");
            CreateTemplate(Path.Combine(root, "YadgTemplates", "mermaid.docx"), includeNumericFigure: true);
            return new(root);
        }
        public void Dispose() { try { Directory.Delete(Root, true); } catch { } }
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Yadg.slnx"))) directory = directory.Parent!;
        return directory?.FullName ?? throw new DirectoryNotFoundException("YADG repository root not found.");
    }

    private static ProcessResult RunProcess(string executable, params string[] arguments)
    {
        var info = new ProcessStartInfo(executable) { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        foreach (var argument in arguments) info.ArgumentList.Add(argument);
        using var process = Process.Start(info)!; process.WaitForExit(); return new(process.ExitCode, process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd());
    }

    private static void CreateTemplate(string path, bool includeNumericFigure = false)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "Caption" }, new Style { Type = StyleValues.Paragraph, StyleId = "Heading1" });
        var body = new Body();
        if (includeNumericFigure)
        {
            body.Append(P("{{yadg:frontmatter}}"), P("version: 1"), P("prototypes:"), P("  figureCaption: fig-prototype"), P("{{/yadg:frontmatter}}"), P("{{yadg:prototype:fig-prototype}}"));
            body.Append(new Paragraph(new Run(new Text("Figure ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), new Run(new FieldCode(" SEQ Figure \\* ARABIC ")), new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), new Run(new Text("0")), new Run(new FieldChar { FieldCharType = FieldCharValues.End }), new Run(new Text(": ")), new Run(new Text("{{caption}}"))));
            body.Append(P("{{/yadg:prototype:fig-prototype}}"), new Paragraph(new Run(new Text("{{content:architecture}}"))), new Paragraph(new Run(new Text("{{figure:system-flow}}"))));
        }
        else body.Append(new Paragraph(new Run(new Text("{{content:architecture}}"))));
        body.Append(new SectionProperties(new PageSize { Width = 12240, Height = 15840 }, new PageMargin { Left = 1440, Right = 1440, Top = 1440, Bottom = 1440 }));
        main.Document = new Document(body); main.Document.Save();
    }

    private static Paragraph P(string text) => new(new Run(new Text(text)));
}
