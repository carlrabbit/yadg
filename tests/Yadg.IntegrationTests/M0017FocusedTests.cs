using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
    }
}
