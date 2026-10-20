using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;
using Yadg.Word;

namespace Yadg.IntegrationTests;

public sealed class M0016FocusedTests
{
    [Fact]
    public void Parser_keeps_authored_fragment_links_separate_from_generated_references()
    {
        var parsed = MarkdownDocumentParser.Parse("# Model {#model}\n\nRead [the model description](#model), [external docs](https://example.test/docs), and [@model].", "content.md");

        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        var paragraph = Assert.Single(parsed.Document!.Blocks.OfType<YadgParagraph>());
        Assert.Contains(paragraph.Inlines.OfType<YadgLink>(), link => link.Destination == "#model" && link.InternalTargetId == "model" && InlineText(link.Inlines) == "the model description");
        Assert.Contains(paragraph.Inlines.OfType<YadgLink>(), link => link.Destination == "https://example.test/docs" && link.InternalTargetId is null && InlineText(link.Inlines) == "external docs");
        Assert.Contains(paragraph.Inlines, inline => inline is YadgReference { Id: "model" });
    }

    [Fact]
    public void Parser_keeps_unsupported_extension_delimiters_literal_instead_of_giving_them_semantics()
    {
        var parsed = MarkdownDocumentParser.Parse("# Model {#model}\n\nThis ~~strikethrough~~ syntax is outside the supported CommonMark/GFM subset.", "content.md");

        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        var paragraph = Assert.Single(parsed.Document!.Blocks.OfType<YadgParagraph>());
        Assert.Contains(paragraph.Inlines.OfType<YadgText>(), text => text.Value.Contains("~~strikethrough~~", StringComparison.Ordinal));
        Assert.DoesNotContain(paragraph.Inlines, inline => inline is YadgEmphasis or YadgStrong);
    }

    [Fact]
    public void Workspace_reference_modes_default_to_numbers_and_reject_unknown_or_invalid_options()
    {
        var diagnostics = new List<Diagnostic>();
        var defaults = WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\n---\n", "YADG.md", diagnostics);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.IsError);
        Assert.Equal(ReferenceFormatting.Default, defaults.References);

        diagnostics.Clear();
        var configured = WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\nreferences:\n  headings: number\n  figures: label-number\n  tables: number\n---\n", "YADG.md", diagnostics);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.IsError);
        Assert.Equal(new ReferenceFormatting("number", "label-number", "number"), configured.References);

        diagnostics.Clear();
        WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\nreferences:\n  headings: label-number\n  figures: page\n  unknown: number\n---\n", "YADG.md", diagnostics);
        Assert.True(diagnostics.Count(diagnostic => diagnostic.IsError) >= 3);
    }

    [Fact]
    public void Word_authoring_emits_semantic_link_bookmarks_and_prototype_owned_label_number_fields()
    {
        using var fixture = Fixture.Create();
        var workspace = WorkspaceLoader.Load(fixture.Root);
        Assert.True(workspace.IsValid, string.Join(Environment.NewLine, workspace.Diagnostics));
        Assert.True(workspace.Document.FindFigure("flow") is not null, string.Join(", ", workspace.Document.Blocks.Select(block => block.GetType().Name)));
        Assert.NotNull(workspace.Document.FindTable("matrix"));
        var analysis = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values);
        Assert.True(analysis.IsValid, string.Join(Environment.NewLine, analysis.Diagnostics));
        var output = Path.Combine(fixture.Root, "authored.docx");
        WordAuthoring.Author(fixture.Template, output, workspace.Document, workspace.Values.Values);

        using var document = WordprocessingDocument.Open(output, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var links = body.Descendants<Hyperlink>().ToArray();
        Assert.Contains(links, link => link.Anchor?.Value?.StartsWith("yadg_intro_", StringComparison.Ordinal) == true && InlineText(link) == "the section overview");
        Assert.Contains(links, link => link.Anchor?.Value?.StartsWith("yadg_flow_", StringComparison.Ordinal) == true && InlineText(link) == "the flow diagram");
        Assert.Contains(links, link => link.Anchor?.Value?.StartsWith("yadg_matrix_", StringComparison.Ordinal) == true && InlineText(link) == "the interface matrix");
        var external = Assert.Single(links, link => link.Id is not null);
        var externalRelationship = document.MainDocumentPart.HyperlinkRelationships.Single(relation => relation.Id == external.Id);
        Assert.Equal("https://example.test/guide", externalRelationship.Uri.ToString());
        Assert.Equal("the external guide", InlineText(external));

        var fields = body.Descendants<FieldCode>().Where(field => field.Text?.Contains("REF", StringComparison.OrdinalIgnoreCase) == true).ToArray();
        Assert.Contains(fields, field => field.Text!.Contains("REF yadg_flow_", StringComparison.Ordinal));
        Assert.Contains(fields, field => field.Text!.Contains("REF yadg_matrix_", StringComparison.Ordinal));
        foreach (var field in fields.Where(field => field.Text!.Contains("figure_", StringComparison.Ordinal) || field.Text.Contains("matrix_", StringComparison.Ordinal)))
        {
            var instruction = field.Text!;
            var bookmark = instruction.Split(' ', StringSplitOptions.RemoveEmptyEntries).First(token => token.StartsWith("yadg_", StringComparison.Ordinal));
            var start = body.Descendants<BookmarkStart>().Single(item => item.Name?.Value == bookmark);
            var end = body.Descendants<BookmarkEnd>().Single(item => item.Id?.Value == start.Id?.Value);
            var paragraph = start.Ancestors<Paragraph>().First();
            var sequence = paragraph.Descendants<FieldCode>().Single(item => item.Text?.Contains("SEQ", StringComparison.OrdinalIgnoreCase) == true);
            var children = paragraph.ChildElements.ToList();
            Assert.True(children.IndexOf(start) < children.IndexOf(sequence.Ancestors<Run>().First()));
            Assert.True(children.IndexOf(sequence.Ancestors<Run>().First()) < children.IndexOf(end));
            Assert.Equal(": ", end.NextSibling<Run>()?.GetFirstChild<Text>()?.Text);
        }
        Assert.Equal(body.Descendants<BookmarkStart>().Count(), body.Descendants<BookmarkEnd>().Count());
    }

    [Fact]
    public void Word_authoring_keeps_legacy_default_references_numeric_only_and_updateable()
    {
        using var fixture = Fixture.Create();
        var workspace = WorkspaceLoader.Load(fixture.Root);
        Assert.True(workspace.IsValid, string.Join(Environment.NewLine, workspace.Diagnostics));
        var legacy = workspace.Document with { ReferenceFormatting = ReferenceFormatting.Default };
        var analysis = WordAuthoring.Analyze(fixture.Template, legacy, workspace.Values.Values);
        Assert.True(analysis.IsValid, string.Join(Environment.NewLine, analysis.Diagnostics));
        var output = Path.Combine(fixture.Root, "legacy.docx");
        WordAuthoring.Author(fixture.Template, output, legacy, workspace.Values.Values);

        using var document = WordprocessingDocument.Open(output, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var refs = body.Descendants<FieldCode>().Where(field => field.Text?.Contains("REF", StringComparison.OrdinalIgnoreCase) == true).ToArray();
        Assert.Contains(refs, field => field.Text!.Contains("REF yadg_flow_", StringComparison.Ordinal) && field.Text!.Contains("\\h", StringComparison.Ordinal));
        Assert.Contains(refs, field => field.Text!.Contains("REF yadg_matrix_", StringComparison.Ordinal) && field.Text!.Contains("\\h", StringComparison.Ordinal));
        foreach (var idPart in new[] { "flow_", "matrix_" })
        {
            var start = Assert.Single(body.Descendants<BookmarkStart>(), bookmark => bookmark.Name?.Value?.Contains(idPart, StringComparison.Ordinal) == true);
            var end = Assert.Single(body.Descendants<BookmarkEnd>(), bookmark => bookmark.Id?.Value == start.Id?.Value);
            var paragraph = start.Ancestors<Paragraph>().First();
            var sequenceRun = paragraph.Descendants<FieldCode>().Single(field => field.Text?.Contains("SEQ", StringComparison.OrdinalIgnoreCase) == true).Ancestors<Run>().First();
            var children = paragraph.ChildElements.ToList();
            Assert.True(children.IndexOf(start) < children.IndexOf(sequenceRun));
            Assert.True(children.IndexOf(sequenceRun) < children.IndexOf(end));
            Assert.Contains(paragraph.Descendants<Text>(), text => text.Text?.Contains(idPart == "flow_" ? "Figure " : "Table ", StringComparison.Ordinal) == true);
        }
    }

    [Fact]
    public void Missing_fragment_targets_fail_strictly_and_degrade_visibly_in_yolo()
    {
        using var fixture = Fixture.Create();
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Intro {#intro}\n\n[missing](#absent).\n");
        var workspace = LoadWorkspace(fixture);

        var strict = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values);
        Assert.False(strict.IsValid);
        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "YADG-LINK-001" && diagnostic.IsError);

        var yolo = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values, yolo: true);
        Assert.True(yolo.IsValid, string.Join(Environment.NewLine, yolo.Diagnostics));
        Assert.Contains(yolo.Diagnostics, diagnostic => diagnostic.Code == "YADG-YOLO-UNRESOLVED-001" && diagnostic.IsDegradation && diagnostic.Message.Contains("preserve authored link text", StringComparison.Ordinal));
        var output = Path.Combine(fixture.Root, "missing-link-yolo.docx");
        WordAuthoring.Author(fixture.Template, output, workspace.Document, workspace.Values.Values, yolo: true);
        using var authored = WordprocessingDocument.Open(output, false);
        var text = string.Concat(authored.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(node => node.Text));
        Assert.Contains("missing", text, StringComparison.Ordinal);
        Assert.DoesNotContain(authored.MainDocumentPart.Document.Body.Descendants<Hyperlink>(), link => InlineText(link) == "missing");
    }

    [Fact]
    public void Nonrendered_fragment_targets_fail_strictly_and_degrade_in_yolo()
    {
        using var fixture = Fixture.Create();
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Intro {#intro}\n\nSee [the flow](#flow).\n\n# Appendix {#appendix}\n\n![Flow diagram](images/flow.png){#flow}\n");
        var workspace = LoadWorkspace(fixture);
        Assert.NotNull(workspace.Document.FindFigure("flow"));

        var strict = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values);
        Assert.False(strict.IsValid);
        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "YADG-LINK-002" && diagnostic.IsError);

        var yolo = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values, yolo: true);
        Assert.True(yolo.IsValid, string.Join(Environment.NewLine, yolo.Diagnostics));
        Assert.Contains(yolo.Diagnostics, diagnostic => diagnostic.Code == "YADG-YOLO-UNRESOLVED-001" && diagnostic.IsDegradation && diagnostic.Message.Contains("not rendered exactly once", StringComparison.Ordinal));
        var output = Path.Combine(fixture.Root, "nonrendered-link-yolo.docx");
        WordAuthoring.Author(fixture.Template, output, workspace.Document, workspace.Values.Values, yolo: true);
        using var authored = WordprocessingDocument.Open(output, false);
        var text = string.Concat(authored.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(node => node.Text));
        Assert.Contains("the flow", text, StringComparison.Ordinal);
        Assert.DoesNotContain(authored.MainDocumentPart.Document.Body.Descendants<Hyperlink>(), link => InlineText(link) == "the flow");
    }

    [Fact]
    public void Duplicate_direct_placements_remain_fatal_in_strict_and_yolo_modes()
    {
        using var fixture = Fixture.Create();
        fixture.AppendBodyText("{{figure:flow}}", "{{figure:flow}}");
        var workspace = LoadWorkspace(fixture);

        var strict = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values);
        Assert.False(strict.IsValid);
        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "YADG-PLACEMENT-002" && diagnostic.IsError);

        var yolo = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values, yolo: true);
        Assert.False(yolo.IsValid);
        Assert.Contains(yolo.Diagnostics, diagnostic => diagnostic.Code == "YADG-PLACEMENT-002" && diagnostic.IsError);
    }

    [Fact]
    public void Invalid_link_destinations_fail_strictly_and_degrade_in_yolo()
    {
        using var fixture = Fixture.Create();
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Intro {#intro}\n\n[invalid destination](http://[).\n");
        var workspace = LoadWorkspace(fixture);

        var strict = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values);
        Assert.False(strict.IsValid);
        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "YADG-LINK-003" && diagnostic.IsError);

        var yolo = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values, yolo: true);
        Assert.True(yolo.IsValid, string.Join(Environment.NewLine, yolo.Diagnostics));
        Assert.Contains(yolo.Diagnostics, diagnostic => diagnostic.Code == "YADG-YOLO-UNRESOLVED-001" && diagnostic.IsDegradation && diagnostic.Message.Contains("http://[", StringComparison.Ordinal));
        var output = Path.Combine(fixture.Root, "invalid-link-yolo.docx");
        WordAuthoring.Author(fixture.Template, output, workspace.Document, workspace.Values.Values, yolo: true);
        using var authored = WordprocessingDocument.Open(output, false);
        var text = string.Concat(authored.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(node => node.Text));
        Assert.Contains("invalid destination", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Ambiguous_caption_prototypes_are_rejected_with_an_actionable_diagnostic()
    {
        using var fixture = Fixture.Create();
        fixture.AppendSecondSequenceToCaptionPrototype();
        var workspace = LoadWorkspace(fixture);

        var strict = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values);
        Assert.False(strict.IsValid);
        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "YADG-REF-006" && diagnostic.IsError && diagnostic.Message.Contains("exactly one SEQ field", StringComparison.Ordinal) && diagnostic.Message.Contains("Update the prototype", StringComparison.Ordinal));

        var yolo = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values, yolo: true);
        Assert.True(yolo.IsValid, string.Join(Environment.NewLine, yolo.Diagnostics));
        Assert.Contains(yolo.Diagnostics, diagnostic => diagnostic.Code == "YADG-YOLO-UNRESOLVED-001" && diagnostic.IsDegradation && diagnostic.Message.Contains("caption prototype", StringComparison.Ordinal));
        var output = Path.Combine(fixture.Root, "ambiguous-caption-yolo.docx");
        WordAuthoring.Author(fixture.Template, output, workspace.Document, workspace.Values.Values, yolo: true);
        using var authored = WordprocessingDocument.Open(output, false);
        var body = authored.MainDocumentPart!.Document.Body!;
        var text = string.Concat(body.Descendants<Text>().Select(node => node.Text));
        Assert.Contains("[@flow]", text, StringComparison.Ordinal);
        Assert.DoesNotContain(body.Descendants<FieldCode>(), field => field.Text?.Contains("REF", StringComparison.OrdinalIgnoreCase) == true && field.Text.Contains("flow", StringComparison.Ordinal));
    }

    private static YadgWorkspace LoadWorkspace(Fixture fixture)
    {
        var workspace = WorkspaceLoader.Load(fixture.Root);
        Assert.True(workspace.IsValid, string.Join(Environment.NewLine, workspace.Diagnostics));
        return workspace;
    }

    private static string InlineText(IEnumerable<YadgInline> inlines) => string.Concat(inlines.Select(inline => inline switch
    {
        YadgText text => text.Value,
        YadgLiteralText text => text.Value,
        YadgEmphasis emphasis => InlineText(emphasis.Inlines),
        YadgStrong strong => InlineText(strong.Inlines),
        _ => string.Empty
    }));

    private static string InlineText(Hyperlink hyperlink) => string.Concat(hyperlink.Descendants<Text>().Select(text => text.Text));

    private sealed class Fixture : IDisposable
    {
        public string Root { get; }
        public string Template => Path.Combine(Root, "YadgTemplates", "report.docx");
        private Fixture(string root) => Root = root;

        public static Fixture Create()
        {
            var root = Path.Combine(Path.GetTempPath(), "yadg-m0016-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "YadgTemplates"));
            Directory.CreateDirectory(Path.Combine(root, "images"));
            File.WriteAllText(Path.Combine(root, "YADG.md"), "---\nyadg:\n  version: 1\nreferences:\n  headings: number\n  figures: label-number\n  tables: label-number\n---\n");
            File.WriteAllText(Path.Combine(root, "content.md"), "# Intro {#intro}\n\nRead [the section overview](#intro), [the flow diagram](#flow), [the interface matrix](#matrix), and [the external guide](https://example.test/guide). See [@flow] and [@matrix].\n\n![Flow diagram](images/flow.png){#flow}\n\n| Name | Protocol |\n|---|---|\n| Alpha | HTTPS |\n{#matrix caption=\"Interface matrix\"}\n");
            File.WriteAllBytes(Path.Combine(root, "images", "flow.png"), Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
            CreateTemplate(Path.Combine(root, "YadgTemplates", "report.docx"));
            return new(root);
        }

        public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, true); }

        public void AppendBodyText(params string[] values)
        {
            using var document = WordprocessingDocument.Open(Template, true);
            var body = document.MainDocumentPart!.Document.Body!;
            foreach (var value in values) body.Append(new Paragraph(new Run(new Text(value))));
            document.MainDocumentPart.Document.Save();
        }

        public void AppendSecondSequenceToCaptionPrototype()
        {
            using var document = WordprocessingDocument.Open(Template, true);
            var paragraph = document.MainDocumentPart!.Document.Body!.Descendants<Paragraph>().Single(candidate => string.Concat(candidate.Descendants<Text>().Select(text => text.Text)).StartsWith("Figure ", StringComparison.Ordinal) && candidate.Descendants<Text>().Any(text => text.Text?.Contains("{{caption}}", StringComparison.Ordinal) == true));
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            paragraph.Append(new Run(new FieldCode(" SEQ Alternate ")));
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            paragraph.Append(new Run(new Text("0")));
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            document.MainDocumentPart.Document.Save();
        }

        private static void CreateTemplate(string path)
        {
            using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            var main = document.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            var heading = new Style { Type = StyleValues.Paragraph, StyleId = "Heading1", StyleParagraphProperties = new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 1 })) };
            stylesPart.Styles = new Styles(heading, new Style { Type = StyleValues.Paragraph, StyleId = "Caption" }, new Style { Type = StyleValues.Table, StyleId = "TableGrid" });
            stylesPart.Styles.Save();
            var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
            var level = new Level { LevelIndex = 0 };
            level.Append(new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = "%1." });
            numberingPart.Numbering = new Numbering(new AbstractNum(level) { AbstractNumberId = 1 }, new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            numberingPart.Numbering.Save();

            var body = new Body();
            AppendText(body, "{{yadg:frontmatter}}", "version: 1", "prototypes:", "  figureCaption: figure", "  tableCaption: table", "{{/yadg:frontmatter}}");
            AppendText(body, "{{yadg:prototype:figure}}");
            body.Append(CaptionPrototype("Figure"));
            AppendText(body, "{{/yadg:prototype:figure}}", "{{yadg:prototype:table}}");
            body.Append(CaptionPrototype("Table"));
            AppendText(body, "{{/yadg:prototype:table}}", "{{section:intro}}");
            body.Append(new SectionProperties(new PageSize { Width = 12240, Height = 15840 }, new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }));
            main.Document = new Document(body);
            main.Document.Save();
        }

        private static Paragraph CaptionPrototype(string label) => new(
            new Run(new Text(label + " ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode($" SEQ {label} \\* ARABIC ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("0")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
            new Run(new Text(": ")),
            new Run(new Text("{{caption}}")));

        private static void AppendText(Body body, params string[] values)
        {
            foreach (var value in values) body.Append(new Paragraph(new Run(new Text(value))));
        }
    }
}
