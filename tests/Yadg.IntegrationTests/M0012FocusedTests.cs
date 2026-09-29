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
    public void Recursive_source_discovery_prunes_dot_and_all_yadg_directories()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        var templates = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templates);
        File.Copy(fixture.Template, Path.Combine(templates, "report.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n");
        var expected = Path.Combine(fixture.Root, "nested"); Directory.CreateDirectory(expected); File.WriteAllText(Path.Combine(expected, "nested.md"), "ordinary");
        foreach (var name in new[] { ".git", ".execution", "YadgTemplates", "YadgPreWords", "YadgWords", "YadgPdfs" })
        {
            var directory = Path.Combine(fixture.Root, name); Directory.CreateDirectory(directory); File.WriteAllText(Path.Combine(directory, "must-not-scan.md"), "hidden source");
        }
        var loaded = WorkspaceLoader.Load(fixture.Root);
        Assert.DoesNotContain(loaded.Diagnostics, d => d.Code == "YADG-WS-006");
        Assert.Equal(new[] { "content.md", "nested/nested.md" }, loaded.Sources.Select(p => Path.GetRelativePath(fixture.Root, p).Replace('\\', '/')).OrderBy(p => p, StringComparer.Ordinal));
        Assert.Single(loaded.Templates);
    }

    [Fact]
    public void Recursive_source_discovery_skips_linked_child_directory_when_supported()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        var templates = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templates); File.Copy(fixture.Template, Path.Combine(templates, "report.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n");
        var external = Path.Combine(Path.GetTempPath(), "yadg-m0012-linked-dir-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(external);
        var link = Path.Combine(fixture.Root, "linked-source");
        try
        {
            File.WriteAllText(Path.Combine(external, "hidden.md"), "must not be discovered");
            try { Directory.CreateSymbolicLink(link, external); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException) { return; }
            var loaded = WorkspaceLoader.Load(fixture.Root);
            Assert.DoesNotContain(loaded.Diagnostics, d => d.IsError);
            Assert.DoesNotContain(loaded.Sources, source => source.Contains("hidden.md", StringComparison.Ordinal));
        }
        finally
        {
            if (Directory.Exists(link)) Directory.Delete(link);
            if (Directory.Exists(external)) Directory.Delete(external, true);
        }
    }

    [Fact]
    public void Publish_cli_destination_bypasses_authoring_and_external_producers()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        var words = Path.Combine(fixture.Root, "YadgWords"); Directory.CreateDirectory(words);
        File.Copy(fixture.Template, Path.Combine(words, "final.docx"));
        var marker = Path.Combine(fixture.Root, "producer-invoked.txt");
        var commandExe = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
        var escapedMarker = marker.Replace("%", "%%", StringComparison.Ordinal);
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), $"---\nyadg:\n  version: 1\nproducers:\n  mermaid:\n    executable: {commandExe}\n    arguments:\n      - /c\n      - echo invoked>{escapedMarker}\n---\n");
        var invalidMarkdown = "# Introduction {#introduction}\n\n---\n\n```mermaid {#sentinel}\nflowchart LR\nA-->B\n```";
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), invalidMarkdown);
        Assert.Contains(MarkdownDocumentParser.Parse(invalidMarkdown, "content.md").Diagnostics, d => d.IsError && d.Code == "YADG-MD-UNSUPPORTED");
        var destination = Path.Combine(fixture.Root, "Published");
        var originalOut = Console.Out; var originalError = Console.Error; using var stdout = new StringWriter(); using var stderr = new StringWriter();
        try { Console.SetOut(stdout); Console.SetError(stderr); Assert.Equal(0, Program.Main(new[] { "publish", "--workspace", fixture.Root, "--publish-path", destination })); }
        finally { Console.SetOut(originalOut); Console.SetError(originalError); }
        Assert.True(File.Exists(Path.Combine(destination, "final.docx")), stderr.ToString());
        Assert.False(File.Exists(marker), "publish executed the configured Mermaid producer");

        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), $"---\nyadg:\n  version: 1\nproducers:\n  mermaid:\n    executable: {commandExe}\n    arguments:\n      - /c\n      - echo invoked>{escapedMarker}\npublish:\n  path: ./Configured\n---\n");
        var configuredOut = Console.Out; var configuredError = Console.Error; using var configuredStdout = new StringWriter(); using var configuredStderr = new StringWriter();
        try { Console.SetOut(configuredStdout); Console.SetError(configuredStderr); Assert.Equal(0, Program.Main(new[] { "publish", "--workspace", fixture.Root })); }
        finally { Console.SetOut(configuredOut); Console.SetError(configuredError); }
        Assert.True(File.Exists(Path.Combine(fixture.Root, "Configured", "final.docx")), configuredStderr.ToString());
        Assert.False(File.Exists(marker), "configured publish executed the Mermaid producer");
    }

    [Fact]
    public void Publish_configured_destination_reads_only_publish_section()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        var words = Path.Combine(fixture.Root, "YadgWords"); Directory.CreateDirectory(words); File.Copy(fixture.Template, Path.Combine(words, "final.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\nmarkdown:\n  unknown: tolerated-for-publication\npublish:\n  path: ./Configured\n---\n");
        var result = WorkspaceValuesParser.ParsePublishPath(Path.Combine(fixture.Root, "YADG.md"), new List<Diagnostic>());
        Assert.Equal("./Configured", result);
        var published = Publisher.Publish(fixture.Root, null, result);
        Assert.True(published.Success, string.Join(Environment.NewLine, published.Diagnostics));
        Assert.True(File.Exists(Path.Combine(fixture.Root, "Configured", "final.docx")));
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
        var strict = WordAuthoring.Analyze(fixture.Template, model);
        Assert.True(strict.IsValid, string.Join(Environment.NewLine, strict.Diagnostics));
        Assert.DoesNotContain(strict.Diagnostics, d => d.Code == "YADG-WORD-LIST");
        Assert.Contains(WordAuthoring.InspectTemplate(fixture.Template), line => line.Contains("role unordered:", StringComparison.Ordinal) && line.Contains("strict=prototype", StringComparison.Ordinal));
    }

    [Fact]
    public void Non_decimal_ordered_style_is_resolved_and_authored_as_a_real_ordered_sequence()
    {
        using var fixture = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "styles:", "  lists:", "    ordered: Roman List", "{{/yadg:frontmatter}}", "{{content:introduction}}");
        fixture.AddOrderedStyle(NumberFormatValues.UpperRoman);
        var model = Parse("# Introduction {#introduction}\n\n1. first\n2. second");
        var analysis = WordAuthoring.Analyze(fixture.Template, model);
        Assert.True(analysis.IsValid, string.Join(Environment.NewLine, analysis.Diagnostics));
        Assert.True(analysis.Presentation["ordered"].StrictResolved);
        Assert.Equal(PresentationSourceKind.ConfiguredResource, analysis.Presentation["ordered"].SourceKind);
        var yolo = WordAuthoring.Analyze(fixture.Template, model, yolo: true);
        Assert.True(yolo.IsValid, string.Join(Environment.NewLine, yolo.Diagnostics));
        Assert.Equal(analysis.Presentation["ordered"].SourceKind, yolo.Presentation["ordered"].SourceKind);
        var output = Path.Combine(fixture.Root, "roman.docx"); WordAuthoring.Author(fixture.Template, output, model);
        using var authored = WordprocessingDocument.Open(output, false);
        var item = authored.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single(p => Text(p) == "first");
        Assert.Equal("RomanList", item.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Contains(WordAuthoring.InspectTemplate(fixture.Template), line => line.Contains("role ordered:", StringComparison.Ordinal) && line.Contains("strict=", StringComparison.Ordinal) && !line.Contains("strict=unresolved", StringComparison.Ordinal));
    }

    [Fact]
    public void Related_list_candidate_selection_is_deterministic()
    {
        using var fixture = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "styles:", "  lists:", "    unordered: Missing Bullet", "{{/yadg:frontmatter}}", "{{content:introduction}}");
        fixture.AddListStyle("Zulu Bullet", "ZuluBullet", 8);
        fixture.AddListStyle("Alpha Bullet", "AlphaBullet", 7);
        var model = Parse("# Introduction {#introduction}\n\n- one");
        var result = WordAuthoring.Analyze(fixture.Template, model, yolo: true);
        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Diagnostics));
        Assert.Equal(PresentationSourceKind.RelatedTemplateResource, result.Presentation["unordered"].SourceKind);
        Assert.Equal("AlphaBullet", result.Presentation["unordered"].SourceId);
        Assert.Contains(WordAuthoring.InspectTemplate(fixture.Template), line => line.Contains("role unordered:", StringComparison.Ordinal) && line.Contains("yolo candidate: Alpha Bullet", StringComparison.Ordinal));
    }

    [Fact]
    public void Table_caption_prototype_precedes_missing_caption_style_in_validation_authoring_and_inspection()
    {
        using var fixture = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "styles:", "  caption: Missing Caption", "prototypes:", "  tableCaption: table-label", "{{/yadg:frontmatter}}", "{{yadg:prototype:table-label}}", "{{caption}}", "{{/yadg:prototype:table-label}}", "{{content:introduction}}");
        fixture.AddTableCaptionField();
        var model = Parse("# Introduction {#introduction}\n\n| A | B |\n|---|---|\n| x | y |\n{#matrix caption=\"Matrix\"}");
        var analysis = WordAuthoring.Analyze(fixture.Template, model);
        Assert.True(analysis.IsValid, string.Join(Environment.NewLine, analysis.Diagnostics));
        Assert.DoesNotContain(analysis.Diagnostics, d => d.Code == "YADG-WORD-STYLE" && d.Message.Contains("caption", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(PresentationSourceKind.Prototype, analysis.Presentation["tableCaption"].SourceKind);
        var output = Path.Combine(fixture.Root, "caption.docx"); WordAuthoring.Author(fixture.Template, output, model);
        using var authored = WordprocessingDocument.Open(output, false);
        Assert.Contains(authored.MainDocumentPart!.Document.Body!.Descendants<FieldCode>(), f => f.Text?.Contains("SEQ Table", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(WordAuthoring.InspectTemplate(fixture.Template), line => line.Contains("role tableCaption:", StringComparison.Ordinal) && line.Contains("strict=prototype", StringComparison.Ordinal));
    }

    [Fact]
    public void Figure_caption_prototype_precedes_missing_caption_style()
    {
        using var fixture = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "styles:", "  caption: Missing Caption", "prototypes:", "  figureCaption: figure-label", "{{/yadg:frontmatter}}", "{{yadg:prototype:figure-label}}", "{{caption}}", "{{/yadg:prototype:figure-label}}", "{{content:introduction}}");
        fixture.AddFigureCaptionField();
        var templateRoot = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templateRoot); File.Copy(fixture.Template, Path.Combine(templateRoot, "report.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n\n![Synthetic](figure.png){#figure}");
        File.WriteAllBytes(Path.Combine(fixture.Root, "figure.png"), Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
        var loaded = WorkspaceLoader.Load(fixture.Root);
        Assert.True(loaded.IsValid, string.Join(Environment.NewLine, loaded.Diagnostics));
        var analysis = WordAuthoring.Analyze(Path.Combine(templateRoot, "report.docx"), loaded.Document!, loaded.Values.Values);
        Assert.True(analysis.IsValid, string.Join(Environment.NewLine, analysis.Diagnostics));
        Assert.Equal(PresentationSourceKind.Prototype, analysis.Presentation["figureCaption"].SourceKind);
    }

    [Fact]
    public void Inline_code_uses_related_character_style_or_plain_fallback()
    {
        var model = Parse("# Introduction {#introduction}\n\nA `snippet`.", "error", "style");
        using (var related = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "styles:", "  codeInline: Missing Code", "{{/yadg:frontmatter}}", "{{content:introduction}}"))
        {
            related.RemoveStyle("CodeCharacter");
            related.AddCharacterStyle("Source Code");
            var output = Path.Combine(related.Root, "related.docx"); WordAuthoring.Author(related.Template, output, model, yolo: true);
            using var doc = WordprocessingDocument.Open(output, false);
            var code = doc.MainDocumentPart!.Document.Body!.Descendants<Run>().Single(r => string.Concat(r.Descendants<Text>().Select(t => t.Text)) == "snippet");
            Assert.Equal("SourceCode", code.RunProperties?.RunStyle?.Val?.Value);
            Assert.Contains(WordAuthoring.Analyze(related.Template, model, yolo: true).Diagnostics, d => d.Code == "YADG-YOLO-CODE-001" && d.IsDegradation);
        }
        using (var plain = Fixture.Create("{{yadg:frontmatter}}", "version: 1", "styles:", "  codeInline: Missing Code", "{{/yadg:frontmatter}}", "{{content:introduction}}"))
        {
            plain.RemoveStyle("CodeCharacter");
            var output = Path.Combine(plain.Root, "plain.docx"); WordAuthoring.Author(plain.Template, output, model, yolo: true);
            using var doc = WordprocessingDocument.Open(output, false);
            var code = doc.MainDocumentPart!.Document.Body!.Descendants<Run>().Single(r => string.Concat(r.Descendants<Text>().Select(t => t.Text)) == "snippet");
            Assert.Null(code.RunProperties?.RunStyle);
            Assert.Contains(WordAuthoring.InspectTemplate(plain.Template), line => line.Contains("yolo candidate: builtin:plain-v1", StringComparison.Ordinal));
        }
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

        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n\nValid content.");
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n  version: 1\n---\n");
        stdout.GetStringBuilder().Clear(); stderr.GetStringBuilder().Clear();
        try { Console.SetOut(stdout); Console.SetError(stderr); Assert.NotEqual(0, Program.Main(new[] { "check", "--yolo", "--workspace", fixture.Root })); }
        finally { Console.SetOut(originalOut); Console.SetError(originalError); }
        Assert.Contains("YADG-VALUES-006", stderr.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Mermaid_failure_is_visible_under_yolo_but_asset_path_escape_stays_fatal()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        var templates = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templates); File.Copy(fixture.Template, Path.Combine(templates, "report.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\nproducers:\n  mermaid:\n    executable: missing-mermaid-executable\n    arguments: []\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n\n```mermaid {#diagram}\nflowchart LR\nA-->B\n```");
        var originalOut = Console.Out; var originalError = Console.Error; using var stdout = new StringWriter(); using var stderr = new StringWriter();
        try
        {
            Console.SetOut(stdout); Console.SetError(stderr);
            Assert.NotEqual(0, Program.Main(new[] { "check", "--workspace", fixture.Root }));
            var yoloCheck = Program.Main(new[] { "check", "--yolo", "--workspace", fixture.Root });
            Assert.True(yoloCheck == 0, $"YOLO check failed ({yoloCheck}):\n{stdout}\n{stderr}");
            Assert.Equal(0, Program.Main(new[] { "build", "--yolo", "--workspace", fixture.Root }));
        }
        finally { Console.SetOut(originalOut); Console.SetError(originalError); }
        Assert.Contains("YADG-YOLO-PRODUCER-001", stderr.ToString(), StringComparison.Ordinal);
        using (var authored = WordprocessingDocument.Open(Path.Combine(fixture.Root, "YadgPreWords", "report.docx"), false))
            Assert.Contains(authored.MainDocumentPart!.Document.Body!.Descendants<Text>(), text => text.Text?.Contains("Mermaid", StringComparison.OrdinalIgnoreCase) == true && text.Text.Contains("unavailable", StringComparison.OrdinalIgnoreCase));

        var model = Parse("![Escape](../outside.png){#escape}");
        var diagnostics = new List<Diagnostic>();
        var result = AssetValidation.AttachAndValidate(model, fixture.Root, diagnostics, yolo: true);
        Assert.Contains(diagnostics, d => d.IsError && d.Code == "YADG-FIGURE-004");
        Assert.Null(result.FindFigure("escape")?.Asset);
        foreach (var (source, id) in new[] { ("![Remote](https://example.invalid/image.png){#remote}", "remote"), ("![Data](data:image/png;base64,AAAA){#data}", "data") })
        {
            var remoteDiagnostics = new List<Diagnostic>();
            var remoteModel = Parse(source);
            AssetValidation.AttachAndValidate(remoteModel, fixture.Root, remoteDiagnostics, yolo: true);
            Assert.Contains(remoteDiagnostics, d => d.IsError && d.Code == "YADG-FIGURE-003");
        }

        var outside = Path.Combine(Path.GetTempPath(), "yadg-m0012-outside-" + Guid.NewGuid().ToString("N") + ".png");
        var link = Path.Combine(fixture.Root, "linked.png");
        try
        {
            File.WriteAllBytes(outside, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
            try
            {
                File.CreateSymbolicLink(link, outside);
                var linkedModel = Parse("![Linked](linked.png){#linked}"); var linkedDiagnostics = new List<Diagnostic>();
                var linkedResult = AssetValidation.AttachAndValidate(linkedModel, fixture.Root, linkedDiagnostics, yolo: true);
                Assert.DoesNotContain(linkedDiagnostics, d => d.IsError);
                Assert.NotNull(linkedResult.FindFigure("linked")?.Asset);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException)
            {
                // File-link behavior is capability-gated; inability to create a link is not a product prerequisite.
            }
        }
        finally { if (File.Exists(link)) File.Delete(link); if (File.Exists(outside)) File.Delete(outside); }
    }

    [Fact]
    public void Renderer_result_exposes_actual_renderer_identity()
    {
        var word = new Yadg.Renderer.RenderResult(true, Array.Empty<Yadg.Renderer.RendererDiagnostic>(), RuntimeVersion: "16.0", ActualRenderer: "word");
        var libreOffice = new Yadg.Renderer.RenderResult(true, Array.Empty<Yadg.Renderer.RendererDiagnostic>(), RuntimeVersion: "26.8", ActualRenderer: "libreoffice");
        Assert.Equal("word", word.ActualRenderer);
        Assert.Equal("libreoffice", libreOffice.ActualRenderer);
    }

    [Fact]
    public void Structured_word_diagnostic_uses_workspace_relative_path_and_inherited_custom_heading_context()
    {
        using var fixture = Fixture.Create("Risk Model", "{{content:introduction}}");
        fixture.SetCustomHeadingContext();
        fixture.RemoveStyle("Heading2");
        var templates = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templates); File.Copy(fixture.Template, Path.Combine(templates, "report.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n\n## Child {#child}\n\nText.");
        var originalOut = Console.Out; var originalError = Console.Error; using var stdout = new StringWriter(); using var stderr = new StringWriter();
        try { Console.SetOut(stdout); Console.SetError(stderr); Assert.NotEqual(0, Program.Main(new[] { "check", "--workspace", fixture.Root })); }
        finally { Console.SetOut(originalOut); Console.SetError(originalError); }
        var output = stderr.ToString() + stdout.ToString();
        Assert.Contains("YadgTemplates/report.docx", output, StringComparison.Ordinal);
        Assert.Contains("Risk Model", output, StringComparison.Ordinal);
        Assert.Contains("near:", output, StringComparison.Ordinal);
        Assert.DoesNotContain(Path.GetFullPath(fixture.Root), output, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Structured_docx_location_keeps_path_story_context_ordinal_and_excerpt_separate()
    {
        var location = new DocxLocation("C:/workspace/YadgTemplates/report.docx", "footer", new[] { "Risk Model" }, 8, "Missing approval value");
        Assert.Equal("C:/workspace/YadgTemplates/report.docx", location.FilePath);
        Assert.Equal("footer", location.Story);
        Assert.Equal(new[] { "Risk Model" }, location.HeadingContext);
        Assert.Equal(8, location.Ordinal);
        Assert.Contains("YadgTemplates/report.docx", location.Format("YadgTemplates/report.docx"), StringComparison.Ordinal);
        Assert.Contains("footer > \"Risk Model\" > paragraph 8", location.Format("YadgTemplates/report.docx"), StringComparison.Ordinal);
        Assert.Contains("near: \"Missing approval value\"", location.Format("YadgTemplates/report.docx"), StringComparison.Ordinal);
    }

    [Fact]
    public void Footer_value_diagnostic_retains_footer_story_and_searchable_excerpt()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        fixture.AddFooterValueTag("{{value:missing-approval}} explain searchable context");
        var model = Parse("# Introduction {#introduction}\n\nBody.");
        var diagnostic = Assert.Single(WordAuthoring.Analyze(fixture.Template, model).Diagnostics, d => d.Code == "YADG-VALUE-002");
        Assert.NotNull(diagnostic.StructuredLocation);
        Assert.Equal("footer", diagnostic.StructuredLocation!.Story);
        Assert.Equal(fixture.Template, diagnostic.StructuredLocation.FilePath);
        Assert.Contains("missing-approval", diagnostic.StructuredLocation.Excerpt, StringComparison.Ordinal);
    }

    [Fact]
    public void Supported_text_box_value_diagnostic_retains_container_story()
    {
        DirectoryInfo? root = new(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "tests", "fixtures", "m0010", "word-origin.docx"))) root = root.Parent;
        Assert.NotNull(root);
        var template = Path.Combine(root!.FullName, "tests", "fixtures", "m0010", "word-origin.docx");
        Assert.True(File.Exists(template), template);
        var diagnostics = WordAuthoring.Analyze(template, Parse("# Introduction {#introduction}\n\nBody."), new Dictionary<string, string>()).Diagnostics;
        Assert.Contains(diagnostics, d => d.Code == "YADG-VALUE-002" && d.StructuredLocation?.Story.EndsWith("/text box", StringComparison.Ordinal) == true && d.StructuredLocation.Excerpt.Contains("story-value", StringComparison.Ordinal));
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
        Assert.DoesNotContain(authored.MainDocumentPart!.Document.Body!.Descendants<Text>(), t => t.Text?.Contains("---", StringComparison.Ordinal) == true);
        Assert.Contains(authored.MainDocumentPart!.Document.Body!.Descendants<Text>(), t => t.Text?.Contains("{{value:missing-id}}", StringComparison.Ordinal) == true);
        var authoredText = string.Concat(authored.MainDocumentPart.Document.Body.Descendants<Text>().Select(t => t.Text));
        Assert.Contains("code", authoredText, StringComparison.Ordinal);
        Assert.Contains("[@missing-reference]", authoredText, StringComparison.Ordinal);
        Assert.Contains("YADG figure \"ghost-figure\" unavailable", authoredText, StringComparison.Ordinal);
        Assert.Contains("YADG-YOLO-UNRESOLVED-001", stderr.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Yolo_template_skip_replaces_complete_output_set_and_no_usable_template_fails()
    {
        using var fixture = Fixture.Create("{{content:introduction}}");
        var templates = Path.Combine(fixture.Root, "YadgTemplates"); Directory.CreateDirectory(templates);
        File.Copy(fixture.Template, Path.Combine(templates, "a.docx"));
        File.Copy(fixture.Template, Path.Combine(templates, "b.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n\nReady.");
        Assert.Equal(0, Program.Main(new[] { "build", "--workspace", fixture.Root }));
        var output = Path.Combine(fixture.Root, "YadgPreWords");
        Assert.Equal(2, Directory.GetFiles(output, "*.docx").Length);

        var bPath = Path.Combine(templates, "b.docx");
        using (var locked = new FileStream(bPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            Assert.Equal(0, Program.Main(new[] { "check", "--yolo", "--workspace", fixture.Root }));
            Assert.Equal(0, Program.Main(new[] { "build", "--yolo", "--workspace", fixture.Root }));
        }
        File.Delete(bPath);
        Assert.Equal(new[] { "a.docx" }, Directory.GetFiles(output, "*.docx").Select(Path.GetFileName).Order().ToArray());

        var prior = File.ReadAllBytes(Path.Combine(output, "a.docx"));
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# First {#same}\n\n# Second {#same}");
        Assert.NotEqual(0, Program.Main(new[] { "build", "--workspace", fixture.Root }));
        Assert.Equal(prior, File.ReadAllBytes(Path.Combine(output, "a.docx")));

        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Introduction {#introduction}\n\nReady.");
        var aPath = Path.Combine(templates, "a.docx");
        using (var locked = new FileStream(aPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            Assert.NotEqual(0, Program.Main(new[] { "check", "--yolo", "--workspace", fixture.Root }));
            Assert.NotEqual(0, Program.Main(new[] { "build", "--yolo", "--workspace", fixture.Root }));
        }
        Assert.Equal(prior, File.ReadAllBytes(Path.Combine(output, "a.docx")));
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
        public void AddOrderedStyle(NumberFormatValues format)
        {
            using var doc = WordprocessingDocument.Open(Template, true); var main = doc.MainDocumentPart!;
            main.StyleDefinitionsPart!.Styles!.Append(new Style { Type = StyleValues.Paragraph, StyleId = "RomanList", StyleName = new StyleName { Val = "Roman List" }, StyleParagraphProperties = new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 3 })) });
            main.StyleDefinitionsPart.Styles.Save();
            var numbering = main.AddNewPart<NumberingDefinitionsPart>();
            numbering.Numbering = new Numbering(new AbstractNum(new Level(new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = format }, new LevelText { Val = "%1." })) { AbstractNumberId = 3 }, new NumberingInstance(new AbstractNumId { Val = 3 }) { NumberID = 3 });
            numbering.Numbering.Save(); main.Document.Save();
        }
        public void AddListStyle(string name, string id, int numberId)
        {
            using var doc = WordprocessingDocument.Open(Template, true); var main = doc.MainDocumentPart!;
            main.StyleDefinitionsPart!.Styles!.Append(new Style { Type = StyleValues.Paragraph, StyleId = id, StyleName = new StyleName { Val = name }, StyleParagraphProperties = new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = numberId })) });
            main.StyleDefinitionsPart.Styles.Save();
            var part = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering ??= new Numbering();
            part.Numbering.Append(new AbstractNum(new Level(new NumberingFormat { Val = NumberFormatValues.Bullet }, new LevelText { Val = "•" })) { AbstractNumberId = numberId }, new NumberingInstance(new AbstractNumId { Val = numberId }) { NumberID = numberId });
            part.Numbering.Save(); main.Document.Save();
        }
        public void AddCharacterStyle(string name)
        {
            using var doc = WordprocessingDocument.Open(Template, true);
            doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new Style { Type = StyleValues.Character, StyleId = "SourceCode", StyleName = new StyleName { Val = name } });
            doc.MainDocumentPart.StyleDefinitionsPart.Styles.Save();
        }
        public void AddTableCaptionField()
        {
            using var doc = WordprocessingDocument.Open(Template, true);
            var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single(p => Text(p) == "{{caption}}");
            paragraph.RemoveAllChildren<Run>();
            paragraph.Append(new Run(new Text("Figure ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" SEQ Table \\* ARABIC ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("1")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(": {{caption}}")));
            doc.MainDocumentPart.Document.Save();
        }
        public void AddFigureCaptionField()
        {
            using var doc = WordprocessingDocument.Open(Template, true);
            var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single(p => Text(p) == "{{caption}}");
            paragraph.RemoveAllChildren<Run>();
            paragraph.Append(new Run(new Text("Figure ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" SEQ Figure \\* ARABIC ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("1")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(": {{caption}}")));
            doc.MainDocumentPart.Document.Body!.Append(new SectionProperties(new PageSize { Width = 12240, Height = 15840 }, new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }));
            doc.MainDocumentPart.Document.Save();
        }
        public void RemoveStyle(string id)
        {
            using var doc = WordprocessingDocument.Open(Template, true);
            var style = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>().Single(s => s.StyleId?.Value == id);
            style.Remove(); doc.MainDocumentPart.StyleDefinitionsPart.Styles.Save();
        }
        public void SetCustomHeadingContext()
        {
            using var doc = WordprocessingDocument.Open(Template, true); var main = doc.MainDocumentPart!; var styles = main.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "BaseOutline", StyleParagraphProperties = new StyleParagraphProperties(new OutlineLevel { Val = 0 }) });
            styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "RiskHeader", StyleName = new StyleName { Val = "Risk Model" }, BasedOn = new BasedOn { Val = "BaseOutline" } });
            var heading = main.Document.Body!.Elements<Paragraph>().First(); heading.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "RiskHeader" });
            styles.Save(); main.Document.Save();
        }
        public void AddFooterValueTag(string text)
        {
            using var doc = WordprocessingDocument.Open(Template, true); var main = doc.MainDocumentPart!;
            var footerPart = main.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(new Run(new Text(text)))); footerPart.Footer.Save();
            var section = main.Document.Body!.Elements<SectionProperties>().LastOrDefault();
            if (section is null) { section = new SectionProperties(); main.Document.Body.Append(section); }
            section.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footerPart) });
            main.Document.Save();
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
}
