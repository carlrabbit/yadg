using System.Text.RegularExpressions;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Extensions.TaskLists;
using Markdig.Renderers.Html;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Yadg.Core;

public static class MarkdownDocumentParser
{
    private static readonly Regex StableId = new(@"^(?<text>.*?)\s*\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\s*$", RegexOptions.Compiled);
    private static readonly Regex TableId = new("^\\s*\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)(?:\\s+caption=\\\"(?<caption>[^\\\"]*)\\\")?\\}\\s*$", RegexOptions.Compiled);
    private static readonly Regex FigureLine = new("^\\s*!\\[(?<alt>.*?)\\]\\((?<path>[^\\s\\)]+)(?:\\s+\"(?<title>[^\"]*)\")?\\)\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)\\}\\s*$", RegexOptions.Compiled);
    private static readonly Regex MermaidInfo = new("^mermaid\\s+\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)(?:\\s+caption=\"(?<caption>[^\"]*)\")?\\}$", RegexOptions.Compiled);
    private static readonly Regex DiagramInfo = new("^diagram\\s+\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)(?:\\s+caption=\"(?<caption>[^\"]*)\")?\\}$", RegexOptions.Compiled);
    private static readonly Regex TextDiagramInfo = new("^text\\s+\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)(?:\\s+caption=\"(?<caption>[^\"]*)\")?\\}$", RegexOptions.Compiled);
    private static readonly Regex MermaidOpening = new("^\\s*`{3,}\\s*(?<info>mermaid.*)$", RegexOptions.Compiled);

    public static ParseResult Parse(string markdown, string location = "<input>", string thematicBreakPolicy = "error", string codeInlinePolicy = "style", bool diagramsAutoDetect = false)
    {
        var pipeline = new MarkdownPipelineBuilder().UsePipeTables().UseGridTables().UseGenericAttributes().UseEmphasisExtras(EmphasisExtraOptions.Strikethrough).UseTaskLists().Build();
        var diagnostics = new List<Diagnostic>();
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        var normalizedSource = string.Join('\n', lines);
        var initialAst = Markdown.ToHtml(normalizedSource, new StringWriter(), pipeline);
        var codeLineRanges = initialAst.OfType<CodeBlock>().Select(code => (Start: code.Line, End: LineAt(normalizedSource, code.Span.End))).ToArray();
        var tableAttributeLines = new Dictionary<int, Match>();
        var parseLines = (string[])lines.Clone();
        for (var i = 0; i < lines.Length; i++)
        {
            if (codeLineRanges.Any(range => i >= range.Start && i <= range.End)) continue;
            var match = TableId.Match(lines[i]);
            if (match.Success) { tableAttributeLines[i] = match; parseLines[i] = string.Empty; }
        }
        var parseSource = string.Join('\n', parseLines);
        // Keep table metadata out of Markdig's row parser, then attach it to the table identified by the AST line span.
        var ast = Markdown.ToHtml(parseSource, new StringWriter(), pipeline);
        var headingMetadata = lines.Select(line => Regex.Match(line, @"^(?<marks>#{1,6})\s+(?<text>.+?)(?:\s*\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)\})?\s*$"))
            .Where(m => m.Success).Select(m => (Text: Regex.Replace(m.Groups["text"].Value.Trim(), @"\s*\{#[A-Za-z][A-Za-z0-9_-]*\}$", ""), Id: m.Groups["id"].Success ? m.Groups["id"].Value : null)).ToArray();
        var figureMetadata = lines.Select(line => FigureLine.Match(line)).Where(m => m.Success)
            .Select(m => (Id: m.Groups["id"].Value, Alt: m.Groups["alt"].Value, Path: m.Groups["path"].Value, HasTitle: m.Groups["title"].Success)).ToArray();
        var mermaidMetadata = lines.Select(line => MermaidOpening.Match(line)).Where(m => m.Success).Select(m =>
        {
            var match = MermaidInfo.Match(m.Groups["info"].Value.Trim());
            if (!match.Success) { diagnostics.Add(new("YADG-MERMAID-001", "Mermaid fence requires exactly a stable ID and optional caption attribute.", true, location)); return (Id: (string?)null, Caption: (string?)null); }
            return (Id: (string?)match.Groups["id"].Value, Caption: match.Groups["caption"].Success ? match.Groups["caption"].Value : null);
        }).ToArray();
        var blocks = new List<YadgBlock>();
        var headingOrdinal = 0; var figureOrdinal = 0; var mermaidOrdinal = 0;
        var astBlocks = ast.ToArray();
        var attachedTableAttributeLines = new HashSet<int>();
        for (var astIndex = 0; astIndex < astBlocks.Length; astIndex++)
        {
            var block = astBlocks[astIndex];
            if (block is Table table)
            {
                var tableEnd = table.OfType<TableRow>().SelectMany(row => row.OfType<TableCell>()).Select(cell => cell.Span.End).DefaultIfEmpty(table.Span.End).Max();
                tableEnd = Math.Max(tableEnd, table.Span.End);
                var lastCellLine = table.OfType<TableRow>().SelectMany(row => row.OfType<TableCell>()).Select(cell => cell.Line).DefaultIfEmpty(LineAt(parseSource, tableEnd)).Max();
                var tableEndLine = Math.Max(lastCellLine, LineAt(parseSource, tableEnd));
                var gridTable = table.OfType<TableRow>().SelectMany(row => row.OfType<TableCell>()).Any(cell => cell.ColumnIndex >= 0);
                var attrLineIndex = tableEndLine + (gridTable ? 2 : 1); // GridTable spans omit its closing separator line.
                var attrMatch = tableAttributeLines.TryGetValue(attrLineIndex, out var matchedAttribute) ? matchedAttribute : Match.Empty;
                if (attrMatch.Success) attachedTableAttributeLines.Add(attrLineIndex);
                var attrs = table.TryGetAttributes();
                var id = attrs?.Id;
                var caption = attrs?.Properties?.FirstOrDefault(p => p.Key == "caption").Value as string;
                if (attrMatch.Success) { id ??= attrMatch.Groups["id"].Value; caption ??= attrMatch.Groups["caption"].Success ? attrMatch.Groups["caption"].Value : null; }
                var convertedTable = ConvertTable(table, id, caption, diagnostics, location, codeInlinePolicy);
                if (convertedTable is not null) blocks.Add(convertedTable);
                continue;
            }
            var converted = ConvertBlock(block, diagnostics, location, headingMetadata, ref headingOrdinal, figureMetadata, ref figureOrdinal, mermaidMetadata, ref mermaidOrdinal, thematicBreakPolicy, codeInlinePolicy, diagramsAutoDetect, lines);
            if (converted is not null) blocks.Add(converted);
        }
        foreach (var attributeLine in tableAttributeLines.Keys.Except(attachedTableAttributeLines)) diagnostics.Add(new("YADG-TABLE-005", "Table attribute line must immediately follow exactly one parsed Markdig table.", true, $"{location}:{attributeLine + 1}:1"));
        if (figureMetadata.Any(f => f.HasTitle)) diagnostics.Add(new("YADG-FIGURE-001", "Markdown image titles are unsupported for M0003 figures.", true, location));
        var sections = new Dictionary<string, YadgSection>(StringComparer.Ordinal);
        var tables = new Dictionary<string, YadgTable>(StringComparer.Ordinal);
        var figures = new Dictionary<string, YadgFigure>(StringComparer.Ordinal);
        for (var i = 0; i < blocks.Count; i++)
        {
            if (blocks[i] is YadgHeading { Id: not null } heading)
            {
                var end = blocks.Count;
                for (var j = i + 1; j < blocks.Count; j++) if (blocks[j] is YadgHeading next && next.Level <= heading.Level) { end = j; break; }
                AddReference(sections, heading.Id!, new YadgSection(heading.Id!, heading, blocks[(i + 1)..end]), diagnostics, location);
            }
            if (blocks[i] is YadgTable table) AddReference(tables, table.Id, table, diagnostics, location);
            if (blocks[i] is YadgFigure figure) AddReference(figures, figure.Id, figure, diagnostics, location);
        }
        return new(new YadgDocument(blocks, sections, tables, figures), diagnostics);
    }

    public static YadgDocument Merge(IEnumerable<(string Location, ParseResult Result)> sources, List<Diagnostic> diagnostics)
    {
        var allBlocks = new List<YadgBlock>();
        var sections = new Dictionary<string, YadgSection>(StringComparer.Ordinal);
        var tables = new Dictionary<string, YadgTable>(StringComparer.Ordinal);
        var figures = new Dictionary<string, YadgFigure>(StringComparer.Ordinal);
        foreach (var (location, result) in sources.OrderBy(s => s.Location, StringComparer.Ordinal))
        {
            diagnostics.AddRange(result.Diagnostics);
            if (result.Document is null) continue;
            allBlocks.AddRange(result.Document.Blocks);
            foreach (var item in result.Document.References) AddGlobal(item.Key, item.Value, sections, tables, figures, diagnostics, location);
            foreach (var item in result.Document.Tables) AddGlobal(item.Key, item.Value, sections, tables, figures, diagnostics, location);
            foreach (var item in result.Document.Figures) AddGlobal(item.Key, item.Value, sections, tables, figures, diagnostics, location);
        }
        return new(allBlocks, sections, tables, figures);
    }

    private static void AddGlobal(string id, object value, Dictionary<string, YadgSection> sections, Dictionary<string, YadgTable> tables, Dictionary<string, YadgFigure> figures, List<Diagnostic> diagnostics, string location)
    {
        if (sections.ContainsKey(id) || tables.ContainsKey(id) || figures.ContainsKey(id)) { diagnostics.Add(new("YADG-REF-002", $"Duplicate stable ID '{id}' across workspace object types.", true, location)); return; }
        switch (value) { case YadgSection section: sections.Add(id, section); break; case YadgTable table: tables.Add(id, table); break; case YadgFigure figure: figures.Add(id, figure); break; }
    }

    private static void AddReference<T>(Dictionary<string, T> target, string id, T value, List<Diagnostic> diagnostics, string location) where T : class
    {
        if (!target.TryAdd(id, value)) diagnostics.Add(new("YADG-REF-002", $"Duplicate stable ID '{id}'.", true, location));
    }

    private static YadgBlock? ConvertBlock(Block block, List<Diagnostic> diagnostics, string location, IReadOnlyList<(string Text, string? Id)> headings, ref int headingOrdinal, IReadOnlyList<(string Id, string Alt, string Path, bool HasTitle)> figures, ref int figureOrdinal, IReadOnlyList<(string? Id, string? Caption)> mermaidMetadata, ref int mermaidOrdinal, string thematicBreakPolicy, string codeInlinePolicy, bool diagramsAutoDetect, string[] sourceLines)
    {
        if (block is ThematicBreakBlock)
        {
            if (thematicBreakPolicy == "error")
            {
                var line = Array.FindIndex(sourceLines, value => Regex.IsMatch(value, @"^\s*(?:(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,})$") );
                var position = line >= 0 ? $"{location}:{line + 1}:{sourceLines[line].IndexOfAny(new[] { '*', '-', '_' }) + 1}" : location;
                diagnostics.Add(new("YADG-MD-UNSUPPORTED", "Unsupported Markdown construct 'thematic break'.", true, position));
            }
            return null;
        }
        if (block is LinkReferenceDefinitionGroup) return null;
        if (block is HeadingBlock heading)
        {
            var text = InlineText(heading.Inline); var match = StableId.Match(text); var metadata = headingOrdinal < headings.Count ? headings[headingOrdinal] : (null, null); headingOrdinal++;
            return metadata.Item1 is not null ? new YadgHeading(heading.Level, metadata.Item1, metadata.Item2) : match.Success ? new YadgHeading(heading.Level, match.Groups["text"].Value, match.Groups["id"].Value) : new YadgHeading(heading.Level, text);
        }
        if (block is ListBlock list)
            return ConvertList(list, diagnostics, location, codeInlinePolicy, 0);
        if (block is QuoteBlock quote)
        {
            var children = new List<YadgBlock>();
            var nestedHeadingOrdinal = 0;
            foreach (var child in quote)
            {
                if (child is not ParagraphBlock and not QuoteBlock) { diagnostics.Add(Positioned("YADG-QUOTE-001", $"Blockquote content '{child.GetType().Name}' is unsupported; only paragraphs and nested blockquotes are supported.", child, location)); continue; }
                var converted = ConvertBlock(child, diagnostics, location, Array.Empty<(string, string?)>(), ref nestedHeadingOrdinal, figures, ref figureOrdinal, mermaidMetadata, ref mermaidOrdinal, thematicBreakPolicy, codeInlinePolicy, diagramsAutoDetect, sourceLines);
                if (converted is not null) children.Add(converted);
            }
            return new YadgBlockQuote(children);
        }
        if (block is CodeBlock codeBlock and not FencedCodeBlock)
        {
            var text = codeBlock.Lines.ToString();
            return new YadgCodeBlock(text, (codeBlock as FencedCodeBlock)?.Info?.Trim());
        }
        if (block is FencedCodeBlock fenced)
        {
            var openingLine = block.Line >= 0 && block.Line < sourceLines.Length ? sourceLines[block.Line] : string.Empty;
            var opening = Regex.Match(openingLine, @"^\s*(?:`{3,}|~{3,})(?<info>.*)$");
            var info = opening.Success ? opening.Groups["info"].Value.Trim() : fenced.Info?.Trim() ?? string.Empty;
            var token = Regex.Match(info, @"^(?<language>[^\s]+)(?:\s+(?<arguments>.*))?$");
            var language = token.Success ? token.Groups["language"].Value : string.Empty;
            var arguments = token.Success && token.Groups["arguments"].Success ? token.Groups["arguments"].Value.Trim() : string.Empty;
            var isExplicitDiagram = string.Equals(language, "diagram", StringComparison.OrdinalIgnoreCase);
            if (isExplicitDiagram)
            {
                var diagramMetadata = DiagramInfo.Match(info);
                if (!diagramMetadata.Success)
                {
                    diagnostics.Add(Positioned("YADG-DIAGRAM-001", "Diagram fence requires exactly a stable ID and optional caption=\"...\" attribute.", block, location));
                    return null;
                }
                var parsed = TextDiagramParser.Parse(fenced.Lines.ToString());
                if (!parsed.Success)
                {
                    diagnostics.Add(Positioned("YADG-DIAGRAM-002", parsed.Error ?? "Diagram geometry is invalid.", block, location));
                    return null;
                }
                if (parsed.Model!.HasInvalidOcclusion)
                {
                    diagnostics.Add(Positioned("YADG-DIAGRAM-003", "Diagram contains an uncovered or ambiguous missing rectangle edge.", block, location));
                    return null;
                }
                if (parsed.Model.Cells.Count == 0)
                {
                    diagnostics.Add(Positioned("YADG-DIAGRAM-002", "Diagram source contains no connected drawing geometry.", block, location));
                    return null;
                }
                return new YadgFigure(diagramMetadata.Groups["id"].Value, diagramMetadata.Groups["caption"].Value, "<generated-diagram.svg>", location) { GeneratedSvg = parsed.Model.Svg };
            }
            if (diagramsAutoDetect && (language.Length == 0 || language == "text") && (language.Length == 0 ? arguments.Length == 0 : arguments.Length == 0 || TextDiagramInfo.IsMatch(info)))
            {
                string? id = null; string? caption = null;
                if (language == "text" && arguments.Length > 0)
                {
                    var candidateMetadata = TextDiagramInfo.Match(info);
                    if (candidateMetadata.Success) { id = candidateMetadata.Groups["id"].Value; caption = candidateMetadata.Groups["caption"].Success ? candidateMetadata.Groups["caption"].Value : null; }
                    else goto NotDiagramCandidate;
                }
                var parsed = TextDiagramParser.Parse(fenced.Lines.ToString());
                if (parsed.Success && parsed.Model!.HasStrongGeometry)
                {
                    if (id is not null) return new YadgFigure(id, caption ?? string.Empty, "<generated-diagram.svg>", location) { GeneratedSvg = parsed.Model.Svg };
                    return new YadgDiagram(parsed.Model.Svg, location);
                }
            }
            NotDiagramCandidate:
            if (!string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase)) return new YadgCodeBlock(fenced.Lines.ToString(), fenced.Info?.Trim());
            if (mermaidOrdinal >= mermaidMetadata.Count || mermaidMetadata[mermaidOrdinal].Id is null)
            {
                diagnostics.Add(Positioned("YADG-MD-UNSUPPORTED", "Only Mermaid fenced blocks with a stable ID are supported.", block, location));
                return null;
            }
            var metadata = mermaidMetadata[mermaidOrdinal++];
            var source = fenced.Lines.ToString().Trim();
            if (source.Length == 0)
            {
                diagnostics.Add(Positioned("YADG-MERMAID-003", "Mermaid source must not be empty.", block, location));
                return null;
            }
            return new YadgFigure(metadata.Id!, metadata.Caption ?? string.Empty, "<generated-mermaid.png>", location, null, source);
        }
        if (block is ParagraphBlock paragraph)
        {
            if (TableId.IsMatch(InlineText(paragraph.Inline))) return null;
            var image = paragraph.Inline?.OfType<LinkInline>().FirstOrDefault(link => link.IsImage);
            if (image is not null)
            {
                if (figureOrdinal >= figures.Count) diagnostics.Add(Positioned("YADG-FIGURE-002", "Images must use block figure syntax with a stable ID.", block, location));
                else { var figure = figures[figureOrdinal++]; return new YadgFigure(figure.Id, figure.Alt, figure.Path, location); }
                return null;
            }
            return new YadgParagraph(ConvertInlines(paragraph.Inline, diagnostics, location, codeInlinePolicy));
        }
        diagnostics.Add(Positioned("YADG-MD-UNSUPPORTED", $"Unsupported Markdown construct '{block.GetType().Name}'.", block, location)); return null;
    }

    internal static YadgTable? ConvertTable(Table table, string? id, string? caption, List<Diagnostic> diagnostics, string location, string codeInlinePolicy)
    {
        if (string.IsNullOrWhiteSpace(id)) { diagnostics.Add(Positioned("YADG-TABLE-001", "Every pipe or grid table must be immediately followed by a stable-ID attribute line.", table, location)); return null; }
        var rows = table.OfType<TableRow>().ToArray();
        if (rows.Length == 0) { diagnostics.Add(Positioned("YADG-TABLE-002", "Markdown table has no rows.", table, location)); return null; }
        var headerRow = rows.FirstOrDefault(r => r.IsHeader) ?? rows[0];
        var gridHeader = headerRow.OfType<TableCell>().Select(c => ConvertTableCell(c, diagnostics, location, codeInlinePolicy)).ToArray();
        var dataRows = rows.Where(r => r != headerRow).Select(r => (IReadOnlyList<YadgTableCell>)r.OfType<TableCell>().Select(c => ConvertTableCell(c, diagnostics, location, codeInlinePolicy)).ToArray()).ToArray();
        var isGridTable = rows.SelectMany(row => row.OfType<TableCell>()).Any(cell => cell.ColumnIndex >= 0);
        var headerWidth = gridHeader.Select(c => c.StartColumn + c.ColumnSpan).DefaultIfEmpty().Max();
        var width = isGridTable ? (table.ColumnDefinitions.Count > 0 ? table.ColumnDefinitions.Count : headerWidth) : gridHeader.Length;
        if (isGridTable && !HasCompleteGridCoverage(rows, width))
        {
            diagnostics.Add(Positioned("YADG-TABLE-003", "Grid table cells do not provide complete, non-overlapping rectangular coverage across all rows and merged spans. Correct the grid borders/cell spans so every logical position is represented exactly once.", table, location));
            return null;
        }
        var legacyHeader = gridHeader.Select(c => c.Paragraphs.FirstOrDefault() ?? Array.Empty<YadgInline>()).ToArray();
        var legacyRows = dataRows.Select(row => (IReadOnlyList<IReadOnlyList<YadgInline>>)row.Select(c => c.Paragraphs.FirstOrDefault() ?? Array.Empty<YadgInline>()).ToArray()).ToArray();
        return new YadgTable(id, legacyHeader, legacyRows, caption) { GridHeader = isGridTable ? gridHeader : null, GridRows = isGridTable ? dataRows : null, IsGridTable = isGridTable };
    }

    internal static bool HasCompleteGridCoverage(IReadOnlyList<TableRow> rows, int width)
    {
        if (width <= 0 || rows.Count == 0) return false;
        var covered = new bool[rows.Count, width];
        foreach (var (row, rowIndex) in rows.Select((value, index) => (value, index)))
        {
            foreach (var cell in row.OfType<TableCell>())
            {
                var startColumn = cell.ColumnIndex;
                var columnSpan = cell.ColumnSpan;
                var rowSpan = cell.RowSpan;
                if (startColumn < 0 || columnSpan < 1 || rowSpan < 1 || startColumn + columnSpan > width || rowIndex + rowSpan > rows.Count)
                    return false;
                for (var r = rowIndex; r < rowIndex + rowSpan; r++)
                    for (var c = startColumn; c < startColumn + columnSpan; c++)
                    {
                        if (covered[r, c]) return false;
                        covered[r, c] = true;
                    }
            }
        }
        for (var r = 0; r < rows.Count; r++)
            for (var c = 0; c < width; c++)
                if (!covered[r, c]) return false;
        return true;
    }

    private static int LineAt(string source, int offset)
    {
        if (source.Length == 0 || offset < 0) return 0;
        var end = Math.Min(offset, source.Length - 1); var line = 0;
        for (var i = 0; i <= end; i++) if (source[i] == '\n') line++;
        return line;
    }

    private static YadgTableCell ConvertTableCell(TableCell cell, List<Diagnostic> diagnostics, string location, string codeInlinePolicy)
    {
        var paragraphs = new List<IReadOnlyList<YadgInline>>();
        foreach (var block in cell)
        {
            if (block is ParagraphBlock paragraph) paragraphs.Add(ConvertInlines(paragraph.Inline, diagnostics, location, codeInlinePolicy));
            else diagnostics.Add(Positioned("YADG-TABLE-004", $"Grid table cell contains unsupported block '{block.GetType().Name}'. Lists, headings, nested tables and code blocks are not supported in table cells.", block, location));
        }
        if (paragraphs.Count == 0) paragraphs.Add(Array.Empty<YadgInline>());
        return new YadgTableCell(paragraphs, Math.Max(1, cell.ColumnSpan), Math.Max(1, cell.RowSpan), cell.ColumnIndex);
    }

    private static YadgList ConvertList(ListBlock list, List<Diagnostic> diagnostics, string location, string codeInlinePolicy, int level)
    {
        var items = new List<YadgListItem>();
        foreach (var item in list.OfType<ListItemBlock>())
        {
            var paragraphs = item.OfType<ParagraphBlock>().ToArray();
            var nested = item.OfType<ListBlock>().ToArray();
            if (paragraphs.Length != 1 || item.Any(child => child is not ParagraphBlock && child is not ListBlock))
            { diagnostics.Add(Positioned("YADG-LIST-001", "List items support one paragraph and nested lists; other multi-paragraph or block content is unsupported.", item, location)); continue; }
            var rawInlines = paragraphs[0].Inline;
            bool? checkedValue = null;
            if (rawInlines is not null)
            {
                var task = rawInlines.OfType<TaskList>().FirstOrDefault();
                if (task is not null)
                {
                    var prop = task.GetType().GetProperty("Checked") ?? task.GetType().GetProperty("IsChecked");
                    checkedValue = prop?.GetValue(task) as bool? ?? false;
                    task.Remove();
                }
            }
            if (level > 8) diagnostics.Add(Positioned("YADG-LIST-002", "List nesting exceeds Word's supported nine effective levels.", item, location));
            items.Add(new YadgListItem(ConvertInlines(rawInlines, diagnostics, location, codeInlinePolicy), nested.Select(n => ConvertList(n, diagnostics, location, codeInlinePolicy, level + 1)).ToArray(), checkedValue));
        }
        return new YadgList(list.IsOrdered, items, level);
    }


    private static IReadOnlyList<YadgInline> ConvertInlines(ContainerInline? container, List<Diagnostic> diagnostics, string location, string codeInlinePolicy = "error")
    {
        var result = new List<YadgInline>(); if (container is null) return result;
        foreach (var inline in container)
        {
            if (inline.GetType().Name == "PipeTableDelimiterInline") continue;
            switch (inline)
            {
                case LiteralInline literal:
                    AppendLiteral(result, literal.Content.ToString());
                    break;
                case LineBreakInline lineBreak: result.Add(lineBreak.IsHard ? new YadgHardBreak() : new YadgSoftBreak()); break;
                case EmphasisInline emphasis:
                    var children = ConvertInlines(emphasis, diagnostics, location, codeInlinePolicy);
                    result.Add(emphasis.GetType().Name.Contains("Strikethrough", StringComparison.Ordinal) || emphasis.DelimiterChar == '~' ? new YadgStrike(children) : emphasis.DelimiterCount >= 2 ? new YadgStrong(children) : new YadgEmphasis(children)); break;
                case LinkInline link when !link.IsImage:
                    var destination = link.Url ?? string.Empty;
                    var internalTarget = destination.StartsWith('#') ? destination[1..] : null;
                    result.Add(new YadgLink(destination, ConvertInlines(link, diagnostics, location, codeInlinePolicy), internalTarget));
                    break;
                case CodeInline code when codeInlinePolicy == "ignore": result.Add(new YadgLiteralText(code.Content)); break;
                case CodeInline code when codeInlinePolicy == "style": result.Add(new YadgCode(code.Content)); break;
                default: diagnostics.Add(Positioned("YADG-MD-UNSUPPORTED", $"Unsupported Markdown inline '{inline.GetType().Name}'.", inline, location)); break;
            }
        }
        for (var i = 0; i < result.Count; i++)
        {
            if (result[i] is not YadgText) continue;
            var combined = (YadgText)result[i]; var j = i + 1;
            while (j < result.Count && result[j] is YadgText next) { combined = new YadgText(combined.Value + next.Value); result[i] = combined; result.RemoveAt(j); }
            var pieces = new List<YadgInline>(); var cursor = 0;
            foreach (Match match in Regex.Matches(combined.Value, @"\[@(?<id>[A-Za-z][A-Za-z0-9_-]*)\]"))
            { if (match.Index > cursor) pieces.Add(new YadgText(combined.Value[cursor..match.Index])); pieces.Add(new YadgReference(match.Groups["id"].Value)); cursor = match.Index + match.Length; }
            if (pieces.Count > 0) { if (cursor < combined.Value.Length) pieces.Add(new YadgText(combined.Value[cursor..])); result.RemoveAt(i); result.InsertRange(i, pieces); i += pieces.Count - 1; }
        }
        return result;
    }

    private static Diagnostic Positioned(string code, string message, MarkdownObject node, string location)
    {
        if (node.Line < 1) return new(code, message, true, location);
        return new(code, message, true, $"{location}:{node.Line + 1}:{Math.Max(1, node.Column + 1)}");
    }

    private static void AppendLiteral(List<YadgInline> result, string value)
    {
        var reference = new Regex(@"\[@(?<id>[A-Za-z][A-Za-z0-9_-]*)\]", RegexOptions.Compiled);
        var cursor = 0;
        foreach (Match match in reference.Matches(value))
        {
            if (match.Index > cursor) AppendTextWithBreaks(result, value[cursor..match.Index]);
            result.Add(new YadgReference(match.Groups["id"].Value));
            cursor = match.Index + match.Length;
        }
        if (cursor < value.Length) AppendTextWithBreaks(result, value[cursor..]);
    }

    private static void AppendTextWithBreaks(List<YadgInline> result, string value)
    {
        var parts = value.Split('\n');
        for (var i = 0; i < parts.Length; i++) { if (parts[i].Length > 0) result.Add(new YadgText(parts[i])); if (i < parts.Length - 1) result.Add(new YadgSoftBreak()); }
    }

    private static string InlineText(ContainerInline? inline) => inline is null ? string.Empty : string.Concat(inline.Select(i => i switch { LiteralInline literal => literal.Content.ToString(), EmphasisInline emphasis => InlineText(emphasis), LineBreakInline => " ", _ => string.Empty }));

    public static (IReadOnlyList<YadgBlock>? Blocks, Diagnostic? Diagnostic) Resolve(YadgDocument document, SectionReference reference, string location = "<template>")
    {
        var section = document.FindSection(reference.Id); return section is null ? (null, new("YADG-REF-001", $"Unresolved stable section ID '{reference.Id}'.", true, location)) : (section.Select(reference.Selection), null);
    }
}
