using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;

namespace Yadg.IntegrationTests;

public sealed class M0006IntegrationTests
{
    [Fact]
    public void Front_matter_is_optional_and_values_keep_a_separate_case_sensitive_namespace()
    {
        using var workspace = TestWorkspace.Create();
        File.WriteAllText(Path.Combine(workspace.Root, "YADG.md"), "---\nyadg:\n  version: 1\nvalues:\n  architecture: \"workspace fact\"\n  Owner: \"upper\"\n  owner: \"lower\"\n---\nThis note is not document source.");
        File.WriteAllText(Path.Combine(workspace.Root, "content.md"), "# Architecture {#architecture}\n\nSee [@architecture].");
        var loaded = WorkspaceLoader.Load(workspace.Root);
        Assert.True(loaded.IsValid, string.Join(Environment.NewLine, loaded.Diagnostics));
        Assert.Equal("workspace fact", loaded.Values.Values["architecture"]);
        Assert.Equal("upper", loaded.Values.Values["Owner"]);
        Assert.Equal("lower", loaded.Values.Values["owner"]);
        Assert.NotNull(loaded.Document.FindSection("architecture"));
        Assert.Contains(loaded.Document.Blocks.OfType<YadgParagraph>().SelectMany(p => p.Inlines), inline => inline is YadgReference reference && reference.Id == "architecture");
        Assert.DoesNotContain(loaded.Document.Blocks.OfType<YadgParagraph>().SelectMany(p => p.Inlines), inline => inline is YadgText text && text.Value.Contains("workspace", StringComparison.Ordinal));

        var noteOnlyDiagnostics = new List<Diagnostic>();
        var noteOnly = WorkspaceValuesParser.Parse("Human-only notes.", "YADG.md", noteOnlyDiagnostics);
        Assert.Empty(noteOnlyDiagnostics);
        Assert.Empty(noteOnly.Values);
    }

    [Fact]
    public void Front_matter_rejects_the_declared_invalid_forms()
    {
        var invalid = new[]
        {
            "---\nyadg:\n  version: 1\nvalues:\n  owner: value",
            "---\nyadg:\n  version: 2\n---",
            "---\nyadg:\n  version: 1\n  extra: x\n---",
            "---\nyadg:\n  version: 1\nvalues:\n  owner: x\n  owner: y\n---",
            "---\nyadg:\n  version: 1\nvalues:\n  1owner: x\n---",
            "---\nyadg:\n  version: 1\nvalues:\n  owner: 42\n---",
            "---\nyadg:\n  version: 1\nvalues:\n  owner: |\n    multiline\n---",
            "---\nyadg: &root\n  version: 1\n---"
        };
        foreach (var source in invalid)
        {
            var diagnostics = new List<Diagnostic>();
            WorkspaceValuesParser.Parse(source, "YADG.md", diagnostics);
            Assert.Contains(diagnostics, diagnostic => diagnostic.IsError);
        }
    }

    [Fact]
    public void Workspace_values_author_all_supported_locations_and_preserve_first_run_formatting()
    {
        using var workspace = TestWorkspace.Create();
        var loaded = WorkspaceLoader.Load(workspace.Root);
        Assert.True(loaded.IsValid, string.Join(Environment.NewLine, loaded.Diagnostics));
        Assert.Equal("2.3", loaded.Values.Values["document-version"]);
        Assert.Equal("", loaded.Values.Values["empty"]);

        var check = RunCli(workspace.Root, "check");
        Assert.Equal(0, check.ExitCode);
        var build = RunCli(workspace.Root, "build");
        Assert.Equal(0, build.ExitCode);

        using var output = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgPreWords", "values.docx"), false);
        var main = output.MainDocumentPart!;
        var all = main.Document.Body!.Descendants<Paragraph>().Concat(main.HeaderParts.SelectMany(p => p.Header.Descendants<Paragraph>())).Concat(main.FooterParts.SelectMany(p => p.Footer.Descendants<Paragraph>())).ToArray();
        var text = all.Select(Logical).ToArray();
        Assert.Contains(text, x => x == "Version Liquidity Risk / 2.3 / {{value:owner}}");
        Assert.Contains(text, x => x == "Cell Liquidity Risk");
        Assert.Contains(text, x => x == "Header 2026-09-30");
        Assert.Contains(text, x => x == "Footer Liquidity Risk");
        Assert.DoesNotContain(text, x => x.Contains("{{value:document-version}}", StringComparison.Ordinal) || x.Contains("{{value:reporting-date}}", StringComparison.Ordinal));
        Assert.Equal(1, text.Sum(x => x.Contains("{{value:owner}}", StringComparison.Ordinal) ? 1 : 0));
        var split = all.Single(p => Logical(p).StartsWith("Version ", StringComparison.Ordinal));
        var ownerRun = split.Descendants<Run>().Single(r => r.Descendants<Text>().Any(t => t.Text?.Contains("Liquidity Risk", StringComparison.Ordinal) == true));
        Assert.NotNull(ownerRun.RunProperties?.Bold);
        Assert.Null(ownerRun.RunProperties?.Italic);
    }

    [Fact]
    public void Invalid_value_front_matter_and_missing_reference_fail_without_changing_output()
    {
        using var workspace = TestWorkspace.Create();
        var output = Path.Combine(workspace.Root, "YadgPreWords", "values.docx");
        File.WriteAllText(output, "sentinel");
        File.WriteAllText(Path.Combine(workspace.Root, "YADG.md"), "---\nyadg:\n  version: 1\nvalues:\n  owner: 42\n---\nNotes");
        var invalid = RunCli(workspace.Root, "check");
        Assert.NotEqual(0, invalid.ExitCode);
        Assert.Contains("YADG-VALUES-011", invalid.Output);
        Assert.Equal("sentinel", File.ReadAllText(output));
    }

    private static string Logical(Paragraph paragraph) => string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
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
        public static TestWorkspace Create()
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0006-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "YadgTemplates")); Directory.CreateDirectory(Path.Combine(root, "YadgPreWords"));
            File.WriteAllText(Path.Combine(root, "YADG.md"), "---\nyadg:\n  version: 1\nvalues:\n  document-version: \"2.3\"\n  reporting-date: \"2026-09-30\"\n  owner: \"Liquidity Risk\"\n  literal: \"{{value:owner}}\"\n  empty: \"\"\n---\n# Human notes\nThis is not source content.");
            File.WriteAllText(Path.Combine(root, "content.md"), "A source paragraph.");
            CreateTemplate(Path.Combine(root, "YadgTemplates", "values.docx"));
            return new(root);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }

    private static void CreateTemplate(string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body(
            new Paragraph(new Run(new Text("Version ")), new Run(new RunProperties(new Bold()), new Text("{{va")), new Run(new RunProperties(new Italic()), new Text("lue:owner}}")), new Run(new Text(" / {{value:document-version}} / {{value:literal}}"))),
            new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("Cell {{value:owner}}")))))),
            new SectionProperties());
        main.Document = new Document(body);
        var header = main.AddNewPart<HeaderPart>(); header.Header = new Header(new Paragraph(new Run(new Text("Header {{value:reporting-date}}")))); header.Header.Save();
        var footer = main.AddNewPart<FooterPart>(); footer.Footer = new Footer(new Paragraph(new Run(new Text("Footer {{value:owner}}")))); footer.Footer.Save();
        body.Elements<SectionProperties>().Single().Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) }, new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
        main.Document.Save();
    }
}
