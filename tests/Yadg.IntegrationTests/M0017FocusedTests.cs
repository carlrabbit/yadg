using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Yadg.Core;
using Yadg.Word;

namespace Yadg.IntegrationTests;

public sealed class M0017FocusedTests
{
    [Fact]
    public void Pipe_tables_use_ast_cells_and_keep_escaped_pipes_code_and_inline_semantics()
    {
        var parsed = MarkdownDocumentParser.Parse("| Name | Detail |\n|---|---|\n| A\\|B | `x|y` and **strong** [link](https://example.test) [@topic] |\n{#matrix caption=\"Table caption\"}", "content.md");

        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        Assert.True(parsed.Document!.Tables.Count > 0, "Parsed block types: " + string.Join(",", parsed.Document.Blocks.Select(b => b.GetType().Name)) + "; diagnostics: " + string.Join(" | ", parsed.Diagnostics));
        var table = Assert.Single(parsed.Document.Tables.Values);
        Assert.Equal("matrix", table.Id);
        Assert.Equal("Table caption", table.Caption);
        Assert.Equal(2, table.Header.Count);
        var first = Assert.Single(table.Rows);
        Assert.Equal(2, first.Count);
        Assert.Contains(first[0].OfType<YadgText>(), text => text.Value.Contains("A|B", StringComparison.Ordinal));
        Assert.Contains(first[1], inline => inline is YadgCode { Value: "x|y" });
        Assert.Contains(first[1].OfType<YadgStrong>(), strong => strong.Inlines.OfType<YadgText>().Any(t => t.Value == "strong"));
        Assert.Contains(first[1], inline => inline is YadgReference { Id: "topic" });
    }

    [Fact]
    public void Grid_tables_preserve_cell_paragraphs_and_report_unsupported_cell_blocks()
    {
        var source = "+-------+-------+\n| Name  | Detail|\n+=======+=======+\n| Alpha | First |\n|       | Second|\n+-------+-------+\n{#grid}";
        var parsed = MarkdownDocumentParser.Parse(source, "grid.md");

        Assert.NotNull(parsed.Document);
        Assert.True(parsed.Document!.Blocks.Any(block => block is YadgTable { Id: "grid" }), "Parsed block types: " + string.Join(",", parsed.Document.Blocks.Select(b => b.GetType().Name)) + "; diagnostics: " + string.Join(" | ", parsed.Diagnostics));
        var table = parsed.Document.Tables["grid"];
        Assert.All(table.GridHeader!, cell => Assert.Single(cell.Paragraphs));
        Assert.Contains(table.GridRows!.SelectMany(row => row).SelectMany(cell => cell.Paragraphs).SelectMany(paragraph => paragraph), inline => inline is YadgSoftBreak or YadgHardBreak);
        Assert.DoesNotContain(parsed.Diagnostics, d => d.Code == "YADG-TABLE-004");
    }

    [Fact]
    public void Grid_tables_preserve_horizontal_and_vertical_merge_spans()
    {
        var source = "+---+---+---+\n| AAAAA | B |\n+---+---+ B +\n| D | E | B |\n+ D +---+---+\n| D | CCCCC |\n+---+---+---+\n{#merge}";
        var parsed = MarkdownDocumentParser.Parse(source, "merge.md");

        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        var table = parsed.Document!.Tables["merge"];
        Assert.Contains(table.GridHeader!, cell => cell.ColumnSpan > 1);
        Assert.Contains(table.GridHeader!, cell => cell.RowSpan > 1);
        Assert.Contains(table.GridRows!.SelectMany(row => row), cell => cell.RowSpan > 1);
    }

    [Fact]
    public void Grid_tables_reject_unrepresentable_geometry_instead_of_filling_structural_gaps()
    {
        const string source = "+---+---+---+\n| A | B | C |\n+===+===+===+\n| D | E | F |\n+---+---+---+";
        var ast = Markdig.Markdown.Parse(source, new MarkdownPipelineBuilder().UseGridTables().Build());
        var table = Assert.Single(ast.OfType<Markdig.Extensions.Tables.Table>());
        var rows = table.OfType<Markdig.Extensions.Tables.TableRow>().ToArray();
        rows[1].Remove(rows[1].OfType<Markdig.Extensions.Tables.TableCell>().Last());
        var diagnostics = new List<Diagnostic>();
        var converted = MarkdownDocumentParser.ConvertTable(table, "bad-grid", null, diagnostics, "bad-grid.md", "style");

        Assert.Null(converted);
        var diagnostic = Assert.Single(diagnostics, item => item.Code == "YADG-TABLE-003");
        Assert.True(diagnostic.IsError);
        Assert.Contains("complete, non-overlapping rectangular coverage", diagnostic.Message);
    }

    [Fact]
    public void Inline_code_defaults_to_template_style_and_unsupported_grid_blocks_are_rejected()
    {
        var values = WorkspaceValuesParser.Parse("---\nyadg:\n  version: 1\n---\n", "YADG.md", new List<Diagnostic>());
        Assert.Equal("style", values.CodeInlinePolicy);

        var source = "+-------+\n| Header|\n+=======+\n| - list|\n+-------+\n{#bad}";
        var parsed = MarkdownDocumentParser.Parse(source, "grid.md");
        Assert.Contains(parsed.Diagnostics, diagnostic => diagnostic.Code == "YADG-TABLE-004");
    }

    [Fact]
    public void Literal_code_nested_lists_quotes_tasks_and_strike_are_semantic_blocks()
    {
        var source = "- outer\n  1. inner\n     - deep\n- [x] done\n- [ ] pending\n\n> quoted **text**\n>\n> > nested quote\n\n```csharp\n  exact()\n{#literal-code-line}\n\n```\n\n    indented()\n\nStruck ~~word~~ and `inline`";
        var parsed = MarkdownDocumentParser.Parse(source, "blocks.md", codeInlinePolicy: "style");

        Assert.True(parsed.IsValid, string.Join(Environment.NewLine, parsed.Diagnostics));
        var lists = parsed.Document!.Blocks.OfType<YadgList>().ToArray();
        Assert.Single(lists);
        Assert.Single(lists[0].Items[0].NestedLists!);
        Assert.Equal(1, Assert.Single(lists[0].Items[0].NestedLists!).Level);
        Assert.Single(Assert.Single(lists[0].Items[0].NestedLists!).Items[0].NestedLists!);
        Assert.Equal(true, lists[0].Items[1].TaskChecked);
        Assert.Equal(false, lists[0].Items[2].TaskChecked);
        Assert.Contains(parsed.Document.Blocks, block => block is YadgBlockQuote);
        var codeBlocks = parsed.Document.Blocks.OfType<YadgCodeBlock>().ToArray();
        Assert.Equal(2, codeBlocks.Length);
        Assert.Contains(codeBlocks, block => block.Text.Contains("  exact()", StringComparison.Ordinal));
        Assert.Contains(codeBlocks, block => block.Text.Contains("{#literal-code-line}", StringComparison.Ordinal));
        Assert.Contains(codeBlocks, block => block.Text.Contains("indented()", StringComparison.Ordinal));
        var reverseNesting = MarkdownDocumentParser.Parse("1. ordered parent\n   - unordered child", "reverse.md");
        var orderedRoot = Assert.Single(reverseNesting.Document!.Blocks.OfType<YadgList>());
        Assert.True(orderedRoot.Ordered);
        Assert.False(Assert.Single(orderedRoot.Items[0].NestedLists!).Ordered);
        var paragraph = parsed.Document.Blocks.OfType<YadgParagraph>().Single();
        Assert.Contains(paragraph.Inlines, inline => inline is YadgStrike);
        Assert.Contains(paragraph.Inlines, inline => inline is YadgCode { Value: "inline" });
    }

    [Fact]
    public void Word_authoring_emits_multiline_cells_spans_strike_and_static_task_glyphs()
    {
        using var fixture = M0016FocusedTests.Fixture.Create();
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Intro {#intro}\n\n- [x] complete\n\n~~removed~~\n\n![Flow diagram](images/flow.png){#flow}\n\n| A | B |\n|---|---|\n| one | two |\n{#matrix caption=\"Matrix\"}\n");
        var workspace = WorkspaceLoader.Load(fixture.Root);
        Assert.True(workspace.IsValid, string.Join(Environment.NewLine, workspace.Diagnostics));
        var analysis = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values, yolo: true);
        Assert.True(analysis.IsValid, string.Join(Environment.NewLine, analysis.Diagnostics));
        var output = Path.Combine(fixture.Root, "m0017.docx");
        WordAuthoring.Author(fixture.Template, output, workspace.Document, workspace.Values.Values, yolo: true);
        using var document = WordprocessingDocument.Open(output, false);
        var body = document.MainDocumentPart!.Document.Body!;
        Assert.Contains(body.Descendants<Text>(), text => text.Text == "☒ ");
        Assert.Contains(body.Descendants<Run>(), run => run.RunProperties?.Strike is not null);
        Assert.Contains(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>(), table => table.Descendants<TableCell>().Any());
    }

    [Fact]
    public void Word_authoring_emits_valid_gridspan_and_vertical_merge_cells()
    {
        using var fixture = M0016FocusedTests.Fixture.Create();
        var source = "# Intro {#intro}\n\n![Flow diagram](images/flow.png){#flow}\n\n+---+---+---+\n| AAAAA | B |\n+---+---+ B +\n| D | E | B |\n+ D +---+---+\n| D | CCCCC |\n+---+---+---+\n{#matrix caption=\"Interface matrix\"}\n";
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), source);
        var workspace = WorkspaceLoader.Load(fixture.Root);
        Assert.True(workspace.IsValid, string.Join(Environment.NewLine, workspace.Diagnostics));
        var analysis = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values);
        Assert.True(analysis.IsValid, string.Join(Environment.NewLine, analysis.Diagnostics));
        var output = Path.Combine(fixture.Root, "grid.docx");
        WordAuthoring.Author(fixture.Template, output, workspace.Document, workspace.Values.Values);
        using var document = WordprocessingDocument.Open(output, false);
        var table = Assert.Single(document.MainDocumentPart!.Document.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>());
        Assert.Contains(table.Descendants<GridSpan>(), span => span.Val?.Value > 1);
        Assert.Contains(table.Descendants<VerticalMerge>(), merge => merge.Val?.Value == MergedCellValues.Restart);
        Assert.Contains(table.Descendants<VerticalMerge>(), merge => merge.Val?.Value == MergedCellValues.Continue);
        Assert.All(table.Elements<TableRow>(), row => Assert.Equal(3, row.Elements<TableCell>().Sum(cell => cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1)));
    }

    [Fact]
    public void Word_authoring_uses_template_blockquote_and_code_styles_and_preserves_nested_levels_and_code_lines()
    {
        using var fixture = M0016FocusedTests.Fixture.Create();
        ConfigureBlockPresentation(fixture.Template);
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Intro {#intro}\n\n> outer **bold**\n>\n> > nested quote\n\n```text\n  first  \n\n    second\n```\n\n        indented first\n\n            indented nested\n");
        var workspace = WorkspaceLoader.Load(fixture.Root);
        Assert.True(workspace.IsValid, string.Join(Environment.NewLine, workspace.Diagnostics));
        var output = Path.Combine(fixture.Root, "block-presentation.docx");
        WordAuthoring.Author(fixture.Template, output, workspace.Document, workspace.Values.Values);

        using var document = WordprocessingDocument.Open(output, false);
        var body = document.MainDocumentPart!.Document.Body!;
        var quoteParagraphs = body.Elements<Paragraph>().Where(paragraph => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "QuotePresentation").ToArray();
        Assert.Equal(2, quoteParagraphs.Length);
        Assert.Contains("outer bold", string.Concat(quoteParagraphs[0].Descendants<Text>().Select(text => text.Text)));
        Assert.Equal("nested quote", string.Concat(quoteParagraphs[1].Descendants<Text>().Select(text => text.Text)));
        Assert.True(int.Parse(quoteParagraphs[1].ParagraphProperties!.Indentation!.Left!.Value!) >
                    int.Parse(quoteParagraphs[0].ParagraphProperties!.Indentation?.Left?.Value ?? "0"));

        var codeParagraphs = body.Elements<Paragraph>().Where(paragraph => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "CodePresentation").ToArray();
        Assert.Equal(new[] { "  first  ", "", "    second", "    indented first", "", "        indented nested" }, codeParagraphs.Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text))));
        Assert.All(codeParagraphs, paragraph => Assert.NotEqual("PlacementNoise", paragraph.ParagraphProperties!.ParagraphStyleId?.Val?.Value));
        Assert.DoesNotContain(body.Elements<Paragraph>(), paragraph => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "PlacementNoise");
    }

    [Fact]
    public void Word_authoring_requires_template_block_presentation_when_quotes_or_code_are_used()
    {
        using var fixture = M0016FocusedTests.Fixture.Create();
        ConfigureBlockPresentation(fixture.Template);
        using (var document = WordprocessingDocument.Open(fixture.Template, true))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Elements<Style>().Single(style => style.StyleId?.Value == "QuotePresentation").Remove();
            styles.Elements<Style>().Single(style => style.StyleId?.Value == "CodePresentation").Remove();
            document.MainDocumentPart.Document.Save();
        }
        File.WriteAllText(Path.Combine(fixture.Root, "content.md"), "# Intro {#intro}\n\n> quoted text\n\n```\ncode\n```\n");
        var workspace = WorkspaceLoader.Load(fixture.Root);
        Assert.True(workspace.IsValid, string.Join(Environment.NewLine, workspace.Diagnostics));
        var analysis = WordAuthoring.Analyze(fixture.Template, workspace.Document, workspace.Values.Values);

        Assert.Contains(analysis.Diagnostics, diagnostic => diagnostic.Code == "YADG-WORD-STYLE" && diagnostic.Message.Contains("QuotePresentation", StringComparison.Ordinal));
        Assert.Contains(analysis.Diagnostics, diagnostic => diagnostic.Code == "YADG-WORD-STYLE" && diagnostic.Message.Contains("CodePresentation", StringComparison.Ordinal));
    }

    private static void ConfigureBlockPresentation(string template)
    {
        using var document = WordprocessingDocument.Open(template, true);
        var main = document.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "QuotePresentation", StyleName = new StyleName { Val = "Quote Presentation" }, StyleParagraphProperties = new StyleParagraphProperties(new Indentation { Left = "720" }) });
        styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "CodePresentation", StyleName = new StyleName { Val = "Code Presentation" }, StyleRunProperties = new StyleRunProperties(new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" }) });
        styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "PlacementNoise", StyleName = new StyleName { Val = "Placement Noise" } });
        var body = main.Document.Body!;
        var closing = body.Elements<Paragraph>().Single(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)).Trim() == "{{/yadg:frontmatter}}");
        closing.InsertBeforeSelf(new Paragraph(new Run(new Text("styles:"))));
        closing.InsertBeforeSelf(new Paragraph(new Run(new Text("  blockquote: QuotePresentation"))));
        closing.InsertBeforeSelf(new Paragraph(new Run(new Text("  codeBlock: CodePresentation"))));
        var placement = body.Elements<Paragraph>().Single(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)).Trim() == "{{section:intro}}");
        placement.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "PlacementNoise" });
        main.Document.Save();
    }
}
