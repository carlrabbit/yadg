using System.Text.RegularExpressions;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Yadg.Core;

public static class MarkdownDocumentParser
{
    private static readonly Regex StableId = new(@"^(?<text>.*?)\s*\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\s*$", RegexOptions.Compiled);
    private static readonly Regex TableId = new("^\\s*\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)(?:\\s+caption=\\\"(?<caption>[^\\\"]*)\\\")?\\}\\s*$", RegexOptions.Compiled);
    private static readonly Regex FigureLine = new("^\\s*!\\[(?<alt>.*?)\\]\\((?<path>[^\\s\\)]+)(?:\\s+\"(?<title>[^\"]*)\")?\\)\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)\\}\\s*$", RegexOptions.Compiled);
    private static readonly Regex MermaidInfo = new("^mermaid\\s+\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)(?:\\s+caption=\"(?<caption>[^\"]*)\")?\\}$", RegexOptions.Compiled);
    private static readonly Regex MermaidOpening = new("^\\s*`{3,}\\s*(?<info>mermaid.*)$", RegexOptions.Compiled);
    private sealed record RawTable(string? Id, string? Caption, string[] Header, string[][] Rows);

    public static ParseResult Parse(string markdown, string location = "<input>")
    {
        var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().UsePipeTables().Build();
        var ast = Markdown.Parse(markdown, pipeline);
        var diagnostics = new List<Diagnostic>();
        var lines = markdown.Replace("\r\n", "\n").Split('\n');
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
        var rawTables = ParseRawTables(lines);
        var blocks = new List<YadgBlock>();
        var headingOrdinal = 0; var tableOrdinal = 0; var figureOrdinal = 0; var mermaidOrdinal = 0;
        foreach (var block in ast)
        {
            var converted = ConvertBlock(block, diagnostics, location, headingMetadata, ref headingOrdinal, rawTables, ref tableOrdinal, figureMetadata, ref figureOrdinal, mermaidMetadata, ref mermaidOrdinal);
            if (converted is not null) blocks.Add(converted);
        }
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

    private static YadgBlock? ConvertBlock(Block block, List<Diagnostic> diagnostics, string location, IReadOnlyList<(string Text, string? Id)> headings, ref int headingOrdinal, IReadOnlyList<RawTable> rawTables, ref int tableOrdinal, IReadOnlyList<(string Id, string Alt, string Path, bool HasTitle)> figures, ref int figureOrdinal, IReadOnlyList<(string? Id, string? Caption)> mermaidMetadata, ref int mermaidOrdinal)
    {
        if (block is LinkReferenceDefinitionGroup) return null;
        if (block is HeadingBlock heading)
        {
            var text = InlineText(heading.Inline); var match = StableId.Match(text); var metadata = headingOrdinal < headings.Count ? headings[headingOrdinal] : (null, null); headingOrdinal++;
            return metadata.Item1 is not null ? new YadgHeading(heading.Level, metadata.Item1, metadata.Item2) : match.Success ? new YadgHeading(heading.Level, match.Groups["text"].Value, match.Groups["id"].Value) : new YadgHeading(heading.Level, text);
        }
        if (block is Table || (block is ParagraphBlock pipeParagraph && pipeParagraph.Inline?.Any(i => i.GetType().Name == "PipeTableDelimiterInline") == true))
        {
            if (tableOrdinal >= rawTables.Count || rawTables[tableOrdinal].Id is null) { diagnostics.Add(new("YADG-TABLE-001", "Every pipe table must be immediately followed by a stable-ID attribute line.", true, location)); return null; }
            var raw = rawTables[tableOrdinal++];
            var header = raw.Header.Select(cell => ParseCell(cell, diagnostics, location)).ToArray();
            var rows = raw.Rows.Select(row => (IReadOnlyList<IReadOnlyList<YadgInline>>)row.Select(cell => ParseCell(cell, diagnostics, location)).ToArray()).ToArray();
            return new YadgTable(raw.Id!, header, rows, raw.Caption);
        }
        if (block is ListBlock list)
        {
            var items = new List<YadgListItem>();
            foreach (var item in list.OfType<ListItemBlock>())
            {
                var paragraphs = item.OfType<ParagraphBlock>().ToArray();
                if (paragraphs.Length != 1 || item.Any(child => child is ListBlock || child is not ParagraphBlock))
                { diagnostics.Add(new("YADG-LIST-001", "Only flat single-paragraph list items are supported.", true, location)); return null; }
                items.Add(new YadgListItem(ConvertInlines(paragraphs[0].Inline, diagnostics, location)));
            }
            return new YadgList(list.IsOrdered, items);
        }
        if (block is FencedCodeBlock fenced)
        {
            if (!string.Equals(fenced.Info, "mermaid", StringComparison.Ordinal) || mermaidOrdinal >= mermaidMetadata.Count || mermaidMetadata[mermaidOrdinal].Id is null)
            {
                diagnostics.Add(new("YADG-MD-UNSUPPORTED", "Only Mermaid fenced blocks with a stable ID are supported.", true, location));
                return null;
            }
            var metadata = mermaidMetadata[mermaidOrdinal++];
            var source = fenced.Lines.ToString().Trim();
            if (source.Length == 0)
            {
                diagnostics.Add(new("YADG-MERMAID-003", "Mermaid source must not be empty.", true, location));
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
                if (figureOrdinal >= figures.Count) diagnostics.Add(new("YADG-FIGURE-002", "Images must use block figure syntax with a stable ID.", true, location));
                else { var figure = figures[figureOrdinal++]; return new YadgFigure(figure.Id, figure.Alt, figure.Path, location); }
                return null;
            }
            return new YadgParagraph(ConvertInlines(paragraph.Inline, diagnostics, location));
        }
        diagnostics.Add(new("YADG-MD-UNSUPPORTED", $"Unsupported Markdown construct '{block.GetType().Name}'.", true, location)); return null;
    }

    private static IReadOnlyList<YadgInline> ConvertCell(TableCell cell, List<Diagnostic> diagnostics, string location)
    {
        var paragraphs = cell.OfType<ParagraphBlock>().ToArray();
        return paragraphs.Length == 1 ? ConvertInlines(paragraphs[0].Inline, diagnostics, location) : Array.Empty<YadgInline>();
    }

    private static IReadOnlyList<YadgInline> ParseCell(string source, List<Diagnostic> diagnostics, string location)
    {
        var paragraph = Markdown.Parse(source, new MarkdownPipelineBuilder().UseAdvancedExtensions().Build()).OfType<ParagraphBlock>().FirstOrDefault();
        return paragraph is null ? Array.Empty<YadgInline>() : ConvertInlines(paragraph.Inline, diagnostics, location);
    }

    private static IReadOnlyList<RawTable> ParseRawTables(string[] lines)
    {
        var tables = new List<RawTable>();
        for (var i = 0; i + 1 < lines.Length; i++)
        {
            if (!IsPipeRow(lines[i]) || !IsSeparatorRow(lines[i + 1])) continue;
            var header = SplitPipeRow(lines[i]); var rows = new List<string[]>(); var j = i + 2;
            while (j < lines.Length && IsPipeRow(lines[j])) { rows.Add(SplitPipeRow(lines[j])); j++; }
            string? id = null; string? caption = null; if (j < lines.Length) { var match = TableId.Match(lines[j]); if (match.Success) { id = match.Groups["id"].Value; caption = match.Groups["caption"].Success ? match.Groups["caption"].Value : null; } }
            tables.Add(new(id, caption, header, rows.ToArray())); i = j;
        }
        return tables;
    }

    private static bool IsPipeRow(string line) => line.Contains('|', StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(line);
    private static bool IsSeparatorRow(string line) => Regex.IsMatch(line, @"^\s*\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)+\|?\s*$");
    private static string[] SplitPipeRow(string line) => line.Trim().Trim('|').Split('|').Select(cell => cell.Trim()).ToArray();

    private static IReadOnlyList<YadgInline> ConvertInlines(ContainerInline? container, List<Diagnostic> diagnostics, string location)
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
                    var children = ConvertInlines(emphasis, diagnostics, location); result.Add(emphasis.DelimiterCount >= 2 ? new YadgStrong(children) : new YadgEmphasis(children)); break;
                default: diagnostics.Add(new("YADG-MD-UNSUPPORTED", $"Unsupported Markdown inline '{inline.GetType().Name}'.", true, location)); break;
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
