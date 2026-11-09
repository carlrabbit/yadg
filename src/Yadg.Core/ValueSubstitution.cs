using System.Text.RegularExpressions;

namespace Yadg.Core;

public static class ValueSubstitution
{
    private static readonly Regex Token = new(@"\{\{value:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}", RegexOptions.Compiled);
    private static readonly Regex ValuePrefix = new(@"\{\{value:", RegexOptions.Compiled);

    public static string Resolve(string text, IReadOnlyDictionary<string, string> values, string location, bool yolo, List<Diagnostic> diagnostics)
    {
        var validStarts = Token.Matches(text).Cast<Match>().Select(match => match.Index).ToHashSet();
        foreach (Match prefix in ValuePrefix.Matches(text))
            if (!validStarts.Contains(prefix.Index)) diagnostics.Add(new("YADG-VALUE-003", "Malformed value token. Expected {{value:<stable-id>}}.", true, location));
        return Token.Replace(text, match =>
        {
            var id = match.Groups["id"].Value;
            if (values.TryGetValue(id, out var value)) return value;
            diagnostics.Add(yolo
                ? new Diagnostic("YADG-YOLO-VALUE-001", $"Workspace value '{id}' is missing; fallback: retain the visible token.", false, location) { IsDegradation = true }
                : new Diagnostic("YADG-VALUE-004", $"Workspace value '{id}' is not defined.", true, location));
            return match.Value;
        });
    }

    public static YadgDocument ApplyToDocument(YadgDocument document, IReadOnlyDictionary<string, string> values, bool yolo, List<Diagnostic> diagnostics, string sourcePath = "Markdown source")
    {
        YadgInline Inline(YadgInline inline, string location) => inline switch
        {
            YadgText text => text with { Value = Resolve(text.Value, values, location, yolo, diagnostics) },
            YadgLiteralText text => text with { Value = Resolve(text.Value, values, location, yolo, diagnostics) },
            YadgCode code => code with { Value = Resolve(code.Value, values, location, yolo, diagnostics) },
            YadgEmphasis emphasis => emphasis with { Inlines = emphasis.Inlines.Select(child => Inline(child, location)).ToArray() },
            YadgStrong strong => strong with { Inlines = strong.Inlines.Select(child => Inline(child, location)).ToArray() },
            YadgStrike strike => strike with { Inlines = strike.Inlines.Select(child => Inline(child, location)).ToArray() },
            YadgLink link => link with { Inlines = link.Inlines.Select(child => Inline(child, location)).ToArray() },
            _ => inline
        };
        YadgBlock Block(YadgBlock block) => block switch
        {
            YadgHeading heading => heading with { Text = Resolve(heading.Text, values, $"{sourcePath}: heading", yolo, diagnostics) },
            YadgParagraph paragraph => paragraph with { Inlines = paragraph.Inlines.Select(inline => Inline(inline, $"{sourcePath}: paragraph")).ToArray() },
            YadgList list => list with { Items = list.Items.Select(item => item with { Inlines = item.Inlines.Select(inline => Inline(inline, $"{sourcePath}: list item")).ToArray(), NestedLists = item.NestedLists?.Select(nested => (YadgList)Block(nested)).ToArray() }).ToArray() },
            YadgTable table => table with
            {
                Caption = table.Caption is null ? null : Resolve(table.Caption, values, $"{sourcePath}: table caption", yolo, diagnostics),
                Header = table.Header.Select(cell => (IReadOnlyList<YadgInline>)cell.Select(inline => Inline(inline, $"{sourcePath}: table cell")).ToArray()).ToArray(),
                Rows = table.Rows.Select(row => (IReadOnlyList<IReadOnlyList<YadgInline>>)row.Select(cell => (IReadOnlyList<YadgInline>)cell.Select(inline => Inline(inline, $"{sourcePath}: table cell")).ToArray()).ToArray()).ToArray(),
                GridHeader = table.GridHeader?.Select(cell => cell with { Paragraphs = cell.Paragraphs.Select(paragraph => (IReadOnlyList<YadgInline>)paragraph.Select(inline => Inline(inline, $"{sourcePath}: grid-table cell")).ToArray()).ToArray() }).ToArray(),
                GridRows = table.GridRows?.Select(row => (IReadOnlyList<YadgTableCell>)row.Select(cell => cell with { Paragraphs = cell.Paragraphs.Select(paragraph => (IReadOnlyList<YadgInline>)paragraph.Select(inline => Inline(inline, $"{sourcePath}: grid-table cell")).ToArray()).ToArray() }).ToArray()).ToArray()
            },
            YadgFigure figure => ResolveFigure(figure),
            YadgCodeBlock code => code with { Text = Resolve(code.Text, values, $"{sourcePath}: code block", yolo, diagnostics) },
            YadgBlockQuote quote => quote with { Blocks = quote.Blocks.Select(Block).ToArray() },
            _ => block
        };
        var resolvedFigures = new Dictionary<string, YadgFigure>(StringComparer.Ordinal);
        YadgFigure ResolveFigure(YadgFigure figure)
        {
            if (resolvedFigures.TryGetValue(figure.Id, out var resolved)) return resolved;
            resolved = figure with { AltText = Resolve(figure.AltText, values, $"{figure.SourcePath}: figure caption '{figure.Id}'", yolo, diagnostics) };
            resolvedFigures.Add(figure.Id, resolved);
            return resolved;
        }
        var blocks = document.Blocks.Select(Block).ToArray();
        var sections = new Dictionary<string, YadgSection>(StringComparer.Ordinal);
        var tables = new Dictionary<string, YadgTable>(StringComparer.Ordinal);
        var figures = document.Figures.ToDictionary(pair => pair.Key, pair => ResolveFigure(pair.Value), StringComparer.Ordinal);
        for (var i = 0; i < blocks.Length; i++)
        {
            if (blocks[i] is YadgHeading { Id: not null } heading)
            {
                var end = blocks.Length;
                for (var j = i + 1; j < blocks.Length; j++) if (blocks[j] is YadgHeading next && next.Level <= heading.Level) { end = j; break; }
                sections[heading.Id!] = new(heading.Id!, heading, blocks[(i + 1)..end]);
            }
            if (blocks[i] is YadgTable table) tables[table.Id] = table;
        }
        return document with { Blocks = blocks, References = sections, Tables = tables, Figures = figures };
    }
}
