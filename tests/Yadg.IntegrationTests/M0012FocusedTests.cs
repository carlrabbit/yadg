using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;
using Yadg.Cli;
using Yadg.Word;

namespace Yadg.IntegrationTests;

public sealed class M0012FocusedTests
{
    [Fact]
    public void Markdown_policies_keep_defaults_and_preserve_ignored_inline_code_text()
    {
        var defaults = MarkdownDocumentParser.Parse("---", "content.md");
        Assert.Contains(defaults.Diagnostics, d => d.Code == "YADG-MD-UNSUPPORTED");
        Assert.Contains(defaults.Diagnostics, d => d.Location == "content.md:1:1");
        var codeError = MarkdownDocumentParser.Parse("first line\n\nA `code` span.", "content.md");
        Assert.Contains(codeError.Diagnostics, d => d.Code == "YADG-MD-UNSUPPORTED" && d.Location is not null && d.Location.StartsWith("content.md:3:", StringComparison.Ordinal));

        var ignored = MarkdownDocumentParser.Parse("before `literal` after\n\n---", "content.md", "ignore", "ignore");
        Assert.DoesNotContain(ignored.Diagnostics, d => d.IsError);
        Assert.Contains(ignored.Document!.Blocks.OfType<YadgParagraph>().SelectMany(p => p.Inlines), i => i is YadgLiteralText { Value: "literal" });
        var codeReference = MarkdownDocumentParser.Parse("`[@not-a-reference]`", "content.md", "error", "ignore");
        Assert.DoesNotContain(codeReference.Document!.Blocks.OfType<YadgParagraph>().SelectMany(p => p.Inlines), inline => inline is YadgReference);
    }

    [Fact]
    public void Workspace_markdown_policy_keys_are_validated_and_default_to_error()
    {
        var diagnostics = new List<Diagnostic>();
        var defaults = WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\n---\n", "YADG.md", diagnostics);
        Assert.Equal("error", defaults.ThematicBreakPolicy);
        Assert.Equal("error", defaults.CodeInlinePolicy);

        diagnostics.Clear();
        var configured = WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\nmarkdown:\n  thematicBreak: ignore\n  codeInline: style\n---\n", "YADG.md", diagnostics);
        Assert.DoesNotContain(diagnostics, d => d.IsError);
        Assert.Equal("ignore", configured.ThematicBreakPolicy);
        Assert.Equal("style", configured.CodeInlinePolicy);

        diagnostics.Clear();
        WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\nmarkdown:\n  codeInline: discard\n---\n", "YADG.md", diagnostics);
        Assert.Contains(diagnostics, d => d.IsError && d.Message.Contains("Unsupported markdown.codeInline mode", StringComparison.Ordinal));
    }

    [Fact]
    public void Init_creates_starter_without_outputs_and_never_overwrites_conflicts()
    {
        var root = Path.Combine(Path.GetTempPath(), "yadg-m0012-init-" + Guid.NewGuid().ToString("N"));
        try
        {
            Assert.Equal(0, Program.Main(new[] { "init", "--workspace", root }));
            Assert.True(File.Exists(Path.Combine(root, "YADG.md")));
            Assert.Contains("introduction", File.ReadAllText(Path.Combine(root, "content.md")), StringComparison.Ordinal);
            Assert.False(Directory.Exists(Path.Combine(root, "YadgPreWords")));
            File.WriteAllText(Path.Combine(root, "content.md"), "keep");
            Assert.Equal(2, Program.Main(new[] { "init", "--workspace", root }));
            Assert.Equal("keep", File.ReadAllText(Path.Combine(root, "content.md")));
            var conflictRoot = Path.Combine(Path.GetTempPath(), "yadg-m0012-conflict-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(conflictRoot);
            try
            {
                File.WriteAllText(Path.Combine(conflictRoot, "content.md"), "unrelated existing owner");
                Assert.Equal(2, Program.Main(new[] { "init", "--workspace", conflictRoot }));
                Assert.False(File.Exists(Path.Combine(conflictRoot, "YADG.md")));
                Assert.False(Directory.Exists(Path.Combine(conflictRoot, "YadgTemplates")));
            }
            finally { Directory.Delete(conflictRoot, true); }
        }
        finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Fact]
    public void Check_list_shows_inventory_and_counts_all_referenceable_objects()
    {
        using var fixture = Fixture.Create("{{content:introduction}}", "");
        Directory.CreateDirectory(Path.Combine(fixture.Root, "YadgTemplates"));
        File.Copy(fixture.Template, Path.Combine(fixture.Root, "YadgTemplates", "report.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n\n| A | B |\n|---|---|\n| x | y |\n{#matrix}");
        var original = Console.Out; using var writer = new StringWriter();
        try
        {
            Console.SetOut(writer);
            Assert.Equal(0, Program.Main(new[] { "check", "--list", "--workspace", fixture.Root }));
        }
        finally { Console.SetOut(original); }
        Assert.Contains("content.md", writer.ToString(), StringComparison.Ordinal);
        Assert.Contains("report.docx", writer.ToString(), StringComparison.Ordinal);
        Assert.Contains("matrix  table", writer.ToString(), StringComparison.Ordinal);
        Assert.Contains("2 reference(s)", writer.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Style_names_and_aliases_resolve_and_styled_inline_code_uses_character_role()
    {
        using var fixture = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "styles:", "  codeInline: Preferred Code Name", "{{/yadg:frontmatter}}", "{{content:introduction}}");
        var model = Parse("# Introduction {#introduction}\n\nA `snippet` here.", "error", "style");
        var output = Path.Combine(fixture.Root, "authored.docx");

        WordAuthoring.Author(fixture.Template, output, model);

        using var document = WordprocessingDocument.Open(output, false);
        var run = document.MainDocumentPart!.Document.Body!.Descendants<Run>().Single(r => string.Concat(r.Descendants<Text>().Select(t => t.Text)) == "snippet");
        Assert.Equal("CodeCharacter", run.RunProperties?.RunStyle?.Val?.Value);
        Assert.Contains(WordAuthoring.InspectStyles(fixture.Template), line => line.Contains("aliases=code-alias", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Style_alias_resolves_to_its_concrete_internal_id()
    {
        using var fixture = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "styles:", "  codeInline: code-alias", "{{/yadg:frontmatter}}", "{{content:introduction}}");
        var model = Parse("# Introduction {#introduction}\n\nA `snippet` here.", "error", "style");
        var output = Path.Combine(fixture.Root, "authored.docx");
        WordAuthoring.Author(fixture.Template, output, model);
        using var document = WordprocessingDocument.Open(output, false);
        var run = document.MainDocumentPart!.Document.Body!.Descendants<Run>().Single(r => string.Concat(r.Descendants<Text>().Select(t => t.Text)) == "snippet");
        Assert.Equal("CodeCharacter", run.RunProperties?.RunStyle?.Val?.Value);
    }

    [Fact]
    public void Unselected_inline_code_does_not_require_a_template_role()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        var model = Parse("# Introduction {#introduction}\n\nPlain selected text.\n\n# Other {#other}\n\nUnselected `code`.", "error", "style");
        var analysis = WordAuthoring.Analyze(fixture.Template, model);
        Assert.DoesNotContain(analysis.Diagnostics, d => d.Code == "YADG-WORD-STYLE" && d.Message.Contains("codeInline", StringComparison.Ordinal));
    }

    [Fact]
    public void List_prototype_is_cloned_with_its_real_numbering_and_item_content()
    {
        using var fixture = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "prototypes:", "  unorderedListItem: bullet-item", "{{/yadg:frontmatter}}", "{{yadg:prototype:bullet-item}}", "{{item}}", "{{/yadg:prototype:bullet-item}}", "{{section:introduction}}");
        fixture.AddNumberedPrototype();
        var model = Parse("# Introduction {#introduction}\n\n- first\n- second references [@target]\n\n## Target {#target}\n\nTarget text.");
        var output = Path.Combine(fixture.Root, "authored.docx");

        WordAuthoring.Author(fixture.Template, output, model);

        using var document = WordprocessingDocument.Open(output, false);
        var items = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Where(p => Text(p).StartsWith("first", StringComparison.Ordinal) || Text(p).StartsWith("second references", StringComparison.Ordinal)).ToArray();
        Assert.Equal(2, items.Length);
        Assert.All(items, p => Assert.Equal(1, p.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value));
        Assert.Contains(items[0].Descendants<Run>(), r => r.RunProperties?.Color?.Val?.Value == "FF0000");
        Assert.Contains(document.MainDocumentPart.Document.Body.Descendants<FieldCode>(), code => code.Text?.Contains("REF", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void Yolo_uses_real_builtin_list_numbering_and_inspect_template_previews_it()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        var model = Parse("# Introduction {#introduction}\n\n- first\n- second");
        var strict = WordAuthoring.Analyze(fixture.Template, model);
        Assert.Contains(strict.Diagnostics, d => d.IsError && d.Code == "YADG-WORD-LIST");
        var bestEffort = WordAuthoring.Analyze(fixture.Template, model, yolo: true);
        Assert.DoesNotContain(bestEffort.Diagnostics, d => d.IsError);
        Assert.Contains(bestEffort.Diagnostics, d => d.IsDegradation && d.Code.StartsWith("YADG-YOLO-LIST", StringComparison.Ordinal));
        var output = Path.Combine(fixture.Root, "best-effort.docx");
        WordAuthoring.Author(fixture.Template, output, model, yolo: true);
        using (var authored = WordprocessingDocument.Open(output, false))
        {
            var rows = authored.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Where(p => Text(p) is "first" or "second").ToArray();
            Assert.Equal(2, rows.Length);
            Assert.All(rows, row => Assert.NotNull(row.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value));
            var numbering = authored.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
            Assert.Contains(numbering.Descendants<NumberingFormat>(), n => n.Val?.Value == NumberFormatValues.Bullet);
        }
        var report = WordAuthoring.InspectTemplate(fixture.Template);
        Assert.Contains(report, line => line.Contains("role unordered", StringComparison.Ordinal) && line.Contains("builtin:unordered-list-v1", StringComparison.Ordinal));
        Assert.Contains(report, line => line.Contains("near:", StringComparison.Ordinal) && line.Contains("content:introduction", StringComparison.Ordinal));

        var templates = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templates); File.Copy(fixture.Template, Path.Combine(templates, "report.docx"));
        var original = Console.Out; using var stdout = new StringWriter();
        try { Console.SetOut(stdout); Assert.Equal(0, Program.Main(new[] { "inspect", "template", "--template", "report.docx", "--workspace", fixture.Root })); }
        finally { Console.SetOut(original); }
        Assert.Contains("frontmatter:", stdout.ToString(), StringComparison.Ordinal);
        Assert.Contains("yolo candidate: builtin:unordered-list-v1", stdout.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Fatal_semantic_and_workspace_errors_remain_errors_with_yolo()
    {
        var parsed = MarkdownDocumentParser.Parse("# First {#same}\n\n# Second {#same}", "content.md");
        Assert.Contains(parsed.Diagnostics, d => d.Code == "YADG-REF-002" && d.IsError);
        var valuesDiagnostics = new List<Diagnostic>();
        WorkspaceValuesParser.Parse("---\nyadg:\n  version: 9\n---\n", "YADG.md", valuesDiagnostics);
        Assert.Contains(valuesDiagnostics, d => d.IsError);

        using var fixture = Fixture.Create("{{content:introduction}}");
        var templates = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templates); File.Copy(fixture.Template, Path.Combine(templates, "report.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# First {#same}\n\n# Second {#same}");
        var originalOut = Console.Out; var originalError = Console.Error; using var stdout = new StringWriter(); using var stderr = new StringWriter();
        try { Console.SetOut(stdout); Console.SetError(stderr); Assert.Equal(2, Program.Main(new[] { "check", "--yolo", "--workspace", fixture.Root })); }
        finally { Console.SetOut(originalOut); Console.SetError(originalError); }
        Assert.Contains("YADG-REF-002", stderr.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Yolo_cli_is_explicit_and_preserves_missing_value_token()
    {
        using var fixture = Fixture.Create("{{content:introduction}}", "{{value:missing-id}}");
        var templates = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templates);
        File.Copy(fixture.Template, Path.Combine(templates, "report.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n\n- item\n\n---\n\nInline `code` and [@missing-reference].\n\n![Ghost alt](missing.png){#ghost-figure}\n");
        var originalOut = Console.Out; var originalError = Console.Error;
        using var stdout = new StringWriter(); using var stderr = new StringWriter();
        try
        {
            Console.SetOut(stdout); Console.SetError(stderr);
            Assert.Equal(2, Program.Main(new[] { "check", "--workspace", fixture.Root }));
            Assert.Equal(0, Program.Main(new[] { "check", "--yolo", "--workspace", fixture.Root }));
            var exit = Program.Main(new[] { "build", "--yolo", "--workspace", fixture.Root });
            Assert.True(exit == 0, $"exit={exit}\nOUT:{stdout}\nERR:{stderr}");
        }
        finally { Console.SetOut(originalOut); Console.SetError(originalError); }
        Assert.Contains("degradations=", stdout.ToString(), StringComparison.Ordinal);
        Assert.Contains("YADG-YOLO-MD-001", stderr.ToString(), StringComparison.Ordinal);
        Assert.Contains("YADG-YOLO-MD-002", stderr.ToString(), StringComparison.Ordinal);
        Assert.Contains("YADG-YOLO-FIGURE-001", stderr.ToString(), StringComparison.Ordinal);
        Assert.Contains("YADG-YOLO-UNRESOLVED-001", stderr.ToString(), StringComparison.Ordinal);
        Assert.True(File.Exists(Path.Combine(fixture.Root, "YadgPreWords", "report.docx")));
        using var authored = WordprocessingDocument.Open(Path.Combine(fixture.Root, "YadgPreWords", "report.docx"), false);
        Assert.Contains(authored.MainDocumentPart!.Document.Body!.Descendants<Text>(), t => t.Text?.Contains("{{value:missing-id}}", StringComparison.Ordinal) == true);
        var authoredText = string.Concat(authored.MainDocumentPart.Document.Body.Descendants<Text>().Select(t => t.Text));
        Assert.Contains("code", authoredText, StringComparison.Ordinal);
        Assert.Contains("[@missing-reference]", authoredText, StringComparison.Ordinal);
        Assert.Contains("YADG figure \"ghost-figure\" unavailable", authoredText, StringComparison.Ordinal);
        Assert.Contains("YADG-YOLO-UNRESOLVED-001", stderr.ToString(), StringComparison.Ordinal);
    }

    private static string Text(Paragraph paragraph) => string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));

    private static YadgDocument Parse(string markdown, string thematic = "error", string code = "error")
    {
        var result = MarkdownDocumentParser.Parse(markdown, "content.md", thematic, code);
        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Diagnostics));
        return result.Document!;
    }

    private sealed class Fixture : IDisposable
    {
        public string Root { get; }
        public string Template { get; }
        private Fixture(string root, string template) { Root = root; Template = template; }
        public static Fixture Create(string first, params string[] paragraphs)
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0012-focused-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            var template = Path.Combine(root, "template.docx");
            using var document = WordprocessingDocument.Create(template, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            var main = document.AddMainDocumentPart();
            var styles = new Styles(
                new Style { Type = StyleValues.Paragraph, StyleId = "Heading1", StyleName = new StyleName { Val = "Heading 1" } },
                new Style { Type = StyleValues.Paragraph, StyleId = "Heading2", StyleName = new StyleName { Val = "Heading 2" } },
                new Style { Type = StyleValues.Paragraph, StyleId = "ListBullet", StyleName = new StyleName { Val = "List Bullet" } },
                new Style { Type = StyleValues.Table, StyleId = "TableGrid", StyleName = new StyleName { Val = "Table Grid" } },
                new Style { Type = StyleValues.Character, StyleId = "CodeCharacter", StyleName = new StyleName { Val = "Preferred Code Name" }, Aliases = new Aliases { Val = "code-alias" } });
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles; main.StyleDefinitionsPart!.Styles!.Save();
            var body = new Body(new Paragraph(new Run(new Text(first)))); foreach (var value in paragraphs) body.Append(new Paragraph(new Run(new Text(value))));
            main.Document = new Document(body); main.Document.Save();
            return new(root, template);
        }
        public void AddNumberedPrototype()
        {
            using var doc = WordprocessingDocument.Open(Template, true); var main = doc.MainDocumentPart!;
            var prototype = main.Document.Body!.Elements<Paragraph>().Single(p => Text(p) == "{{item}}");
            prototype.ParagraphProperties = new ParagraphProperties(new NumberingProperties(new NumberingId { Val = 1 }));
            prototype.RemoveAllChildren<Run>();
            prototype.Append(new Run(new RunProperties(new Color { Val = "FF0000" }), new Text("{{it")));
            prototype.Append(new Run(new Text("em}}")));
            var numbering = main.AddNewPart<NumberingDefinitionsPart>();
            numbering.Numbering = new Numbering(
                new AbstractNum(new MultiLevelType { Val = MultiLevelValues.SingleLevel }, new Level(new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = NumberFormatValues.Bullet }, new LevelText { Val = "•" })) { AbstractNumberId = 0 },
                new AbstractNum(new MultiLevelType { Val = MultiLevelValues.Multilevel }, new Level(new NumberingFormat { Val = NumberFormatValues.Decimal }, new ParagraphStyleIdInLevel { Val = "Heading1" }) { LevelIndex = 0 }, new Level(new NumberingFormat { Val = NumberFormatValues.Decimal }, new ParagraphStyleIdInLevel { Val = "Heading2" }) { LevelIndex = 1 }) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 }, new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 2 });
            numbering.Numbering.Save(); main.Document.Save();
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
}
