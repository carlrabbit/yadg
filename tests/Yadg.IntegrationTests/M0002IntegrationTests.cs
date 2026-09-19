using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;

namespace Yadg.IntegrationTests;

public sealed class M0002IntegrationTests
{
    [Fact]
    public void Workspace_merges_sources_and_excludes_conventional_directories()
    {
        using var workspace = SyntheticWorkspace.Create();
        var loaded = WorkspaceLoader.Load(workspace.Root);
        Assert.True(loaded.IsValid, string.Join(Environment.NewLine, loaded.Diagnostics));
        Assert.Equal(2, loaded.Sources.Count);
        Assert.DoesNotContain(loaded.Sources, path => path.Contains("YadgTemplates", StringComparison.Ordinal));
        Assert.Equal("architecture", loaded.Document.FindSection("architecture")!.Id);
        Assert.Null(loaded.Document.FindSection("Architecture"));
    }

    [Fact]
    public void Structured_parser_preserves_inline_semantics_and_rejects_lists()
    {
        var parsed = MarkdownDocumentParser.Parse("# Heading {#heading}\n\nplain *italic* and **bold**  \nnext");
        Assert.True(parsed.IsValid);
        var paragraph = Assert.IsType<YadgParagraph>(parsed.Document!.Blocks[1]);
        Assert.Contains(paragraph.Inlines, inline => inline is YadgEmphasis);
        Assert.Contains(paragraph.Inlines, inline => inline is YadgStrong);
        Assert.Contains(paragraph.Inlines, inline => inline is YadgHardBreak);
        Assert.Equal("Heading", Assert.IsType<YadgHeading>(parsed.Document.Blocks[0]).Text);

        var unsupported = MarkdownDocumentParser.Parse("- list");
        Assert.Contains(unsupported.Diagnostics, d => d.Code == "YADG-MD-UNSUPPORTED");
    }

    [Fact]
    public void Nested_workspace_marker_is_rejected()
    {
        using var workspace = SyntheticWorkspace.Create();
        File.WriteAllText(Path.Combine(workspace.Root, "nested", "YADG.md"), "nested");
        var loaded = WorkspaceLoader.Load(workspace.Root);
        Assert.Contains(loaded.Diagnostics, diagnostic => diagnostic.Code == "YADG-WS-002");
    }

    [Fact]
    public void Missing_workspace_inputs_and_heading_styles_are_diagnosed()
    {
        var root = Path.Combine(Path.GetTempPath(), "yadg-m0002-missing-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var missing = WorkspaceLoader.Load(root);
            Assert.Contains(missing.Diagnostics, diagnostic => diagnostic.Code == "YADG-WS-001");
            Assert.Contains(missing.Diagnostics, diagnostic => diagnostic.Code == "YADG-WS-003");
            Assert.Contains(missing.Diagnostics, diagnostic => diagnostic.Code == "YADG-WS-004");

            using var workspace = SyntheticWorkspace.Create();
            CreateTemplate(Path.Combine(workspace.Root, "YadgTemplates", "nostyles.docx"), "{{content:architecture}}", includeStyles: false);
            var result = RunCli(workspace.Root, "check");
            Assert.Contains("YADG-WORD-STYLE", result.Output);
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void Real_cli_builds_two_templates_and_preserves_word_structure()
    {
        using var workspace = SyntheticWorkspace.Create();
        var check = RunCli(workspace.Root, "check");
        Assert.True(check.ExitCode == 0, check.Output);
        var explicitCheck = RunCli(workspace.Root, "check", "--workspace", workspace.Root);
        Assert.Equal(0, explicitCheck.ExitCode);
        var build = RunCli(workspace.Root, "build");
        Assert.True(build.ExitCode == 0, build.Output);
        Assert.True(File.Exists(Path.Combine(workspace.Root, "YadgPreWords", "first.docx")));
        Assert.True(File.Exists(Path.Combine(workspace.Root, "YadgPreWords", "second.docx")));
        Assert.True(File.Exists(Path.Combine(workspace.Root, "YadgPreWords", "keep.txt")));

        using var first = WordprocessingDocument.Open(Path.Combine(workspace.Root, "YadgPreWords", "first.docx"), false);
        var paragraphs = first.MainDocumentPart!.Document.Body!.Descendants<Paragraph>().ToArray();
        var texts = paragraphs.Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))).ToArray();
        Assert.Contains(texts, text => text.Contains("Intro", StringComparison.Ordinal));
        Assert.Contains("Template-owned paragraph", texts);
        Assert.DoesNotContain(texts, text => text.Contains("{{content:", StringComparison.Ordinal));
        Assert.Contains(paragraphs.SelectMany(p => p.Descendants<Run>()), run => run.RunProperties?.Italic is not null);
        Assert.Contains(paragraphs.SelectMany(p => p.Descendants<Run>()), run => run.RunProperties?.Bold is not null);
        Assert.NotEmpty(paragraphs.SelectMany(p => p.Descendants<Break>()));
        Assert.Contains(paragraphs, p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Heading3");
    }

    [Fact]
    public void Invalid_workspace_fails_before_outputs_change_and_reports_reserved_value()
    {
        using var workspace = SyntheticWorkspace.Create();
        File.WriteAllText(Path.Combine(workspace.Root, "duplicate.md"), "# Duplicate {#architecture}\n\nBad.");
        var output = Path.Combine(workspace.Root, "YadgPreWords", "first.docx");
        File.WriteAllText(output, "sentinel");
        File.WriteAllText(Path.Combine(workspace.Root, "bad.md"), "# Unsupported {#bad}\n\n[link](https://example.test)");
        var result = RunCli(workspace.Root, "build");
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("YADG-REF-002", result.Output);
        Assert.Equal("sentinel", File.ReadAllText(output));

        File.Delete(Path.Combine(workspace.Root, "duplicate.md"));
        File.Delete(Path.Combine(workspace.Root, "bad.md"));
        File.WriteAllText(Path.Combine(workspace.Root, "value.md"), "# Value {#value}\n\ntext");
        CreateTemplate(Path.Combine(workspace.Root, "YadgTemplates", "value.docx"), "{{value:value}}", includeStyles: true);
        var valueResult = RunCli(workspace.Root, "check");
        Assert.Contains("YADG-VALUE-UNSUPPORTED", valueResult.Output);
    }

    [Fact]
    public void Old_and_mixed_tags_are_rejected()
    {
        using var workspace = SyntheticWorkspace.Create();
        CreateTemplate(Path.Combine(workspace.Root, "YadgTemplates", "old.docx"), "{{yadg:section:content:architecture}}", includeStyles: true);
        var oldResult = RunCli(workspace.Root, "check");
        Assert.Contains("YADG-TAG-001", oldResult.Output);
        CreateTemplate(Path.Combine(workspace.Root, "YadgTemplates", "mixed.docx"), "prefix {{content:architecture}}", includeStyles: true);
        var mixedResult = RunCli(workspace.Root, "check");
        Assert.Contains("YADG-WORD-BLOCK", mixedResult.Output);
        CreateTableTemplate(Path.Combine(workspace.Root, "YadgTemplates", "table.docx"), "{{content:architecture}}");
        var tableResult = RunCli(workspace.Root, "check");
        Assert.Contains("YADG-WORD-LOCATION", tableResult.Output);
    }

    private static ProcessResult RunCli(string root, string command, params string[] extraArguments)
    {
        var executable = Path.Combine(AppContext.BaseDirectory, "yadg.exe");
        var startInfo = new ProcessStartInfo(executable) { WorkingDirectory = root, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        startInfo.ArgumentList.Add(command);
        foreach (var argument in extraArguments) startInfo.ArgumentList.Add(argument);
        var process = Process.Start(startInfo)!;
        process.WaitForExit();
        return new(process.ExitCode, process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd());
    }

    private sealed record ProcessResult(int ExitCode, string Output);

    private sealed class SyntheticWorkspace : IDisposable
    {
        public string Root { get; }
        private SyntheticWorkspace(string root) => Root = root;
        public static SyntheticWorkspace Create()
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0002-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "nested"));
            Directory.CreateDirectory(Path.Combine(root, "YadgTemplates"));
            Directory.CreateDirectory(Path.Combine(root, "YadgPreWords"));
            Directory.CreateDirectory(Path.Combine(root, ".ignored"));
            File.WriteAllText(Path.Combine(root, "YADG.md"), "Synthetic workspace notes.");
            File.WriteAllText(Path.Combine(root, "a.md"), "# Architecture {#architecture}\n\nIntro *italic* and **bold**  \nnext.\n\n### Nested\n\nNested text.");
            File.WriteAllText(Path.Combine(root, "nested", "b.md"), "## Deployment {#deployment}\n\nDeploy text.");
            File.WriteAllText(Path.Combine(root, ".ignored", "ignored.md"), "# Ignored {#ignored}");
            File.WriteAllText(Path.Combine(root, "YadgTemplates", "ignored.md"), "# Not a source");
            File.WriteAllText(Path.Combine(root, "YadgPreWords", "keep.txt"), "keep");
            CreateTemplate(Path.Combine(root, "YadgTemplates", "first.docx"), "{{content:architecture}}", true);
            CreateTemplate(Path.Combine(root, "YadgTemplates", "second.docx"), "{{section:deployment}}", true);
            return new(root);
        }
        public void Dispose() => Directory.Delete(Root, true);
    }

    private static void CreateTemplate(string path, string tag, bool includeStyles)
    {
        using var document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        if (includeStyles)
        {
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "BodyText" });
            stylesPart.Styles.Append(Enumerable.Range(1, 6).Select(i => new Style { Type = StyleValues.Paragraph, StyleId = $"Heading{i}" }));
            stylesPart.Styles.Save();
        }
        main.Document = new Document(new Body(
            new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("Template heading"))),
            new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "BodyText" }), new Run(new Text("{{")), new Run(new Text(tag[2..^2])), new Run(new Text("}}"))),
            new Paragraph(new Run(new Text("Template-owned paragraph")))));
        main.Document.Save();
    }

    private static void CreateTableTemplate(string path, string tag)
    {
        using var document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "BodyText" });
        stylesPart.Styles.Append(Enumerable.Range(1, 6).Select(i => new Style { Type = StyleValues.Paragraph, StyleId = $"Heading{i}" }));
        stylesPart.Styles.Save();
        main.Document = new Document(new Body(new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text(tag))))))));
        main.Document.Save();
    }
}
