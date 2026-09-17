using Markdig;
using System.Collections.ObjectModel;

namespace Yadg.Core;

public static class MarkdownDocumentParser
{
    public static ParseResult Parse(string markdown)
    {
        var diagnostics = new List<Diagnostic>();
        _ = Markdown.Parse(markdown, new MarkdownPipelineBuilder().Build());
        var lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var headings = new List<(int Line, int Level, string Text, string Id)>();
        for (var index = 0; index < lines.Length; index++)
        {
            var match = System.Text.RegularExpressions.Regex.Match(lines[index], "^(?<marks>#{1,6})\\s+(?<text>.+?)\\s*\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)\\}\\s*$");
            if (!match.Success) continue;
            headings.Add((index, match.Groups["marks"].Length, match.Groups["text"].Value.Trim(), match.Groups["id"].Value));
        }

        var duplicate = headings.GroupBy(h => h.Id, StringComparer.Ordinal).Where(g => g.Count() > 1);
        foreach (var group in duplicate) diagnostics.Add(new("YADG-REF-002", $"Duplicate stable section ID '{group.Key}'."));
        var sections = new List<YadgSection>();
        foreach (var heading in headings)
        {
            var end = lines.Length;
            for (var i = heading.Line + 1; i < lines.Length; i++)
            {
                var next = System.Text.RegularExpressions.Regex.Match(lines[i], "^(?<marks>#{1,6})\\s+");
                if (next.Success && next.Groups["marks"].Length <= heading.Level) { end = i; break; }
            }
            var body = lines[(heading.Line + 1)..end]
                .Select(line => line.TrimEnd())
                .SkipWhile(string.IsNullOrWhiteSpace)
                .Reverse().SkipWhile(string.IsNullOrWhiteSpace).Reverse()
                .ToArray();
            sections.Add(new(heading.Id, heading.Text, heading.Level, body));
        }
        if (headings.Count == 0) diagnostics.Add(new("YADG-MD-001", "Markdown contains no heading with an explicit stable ID."));
        return new(new YadgDocument(new ReadOnlyCollection<YadgSection>(sections)), diagnostics);
    }

    public static (string? Content, Diagnostic? Diagnostic) Resolve(YadgDocument document, SectionReference reference)
    {
        var section = document.FindSection(reference.Id);
        return section is null
            ? (null, new("YADG-REF-001", $"Unresolved stable section ID '{reference.Id}'."))
            : (reference.Selection == SectionSelection.Content ? section.ContentText : section.SectionText, null);
    }
}
