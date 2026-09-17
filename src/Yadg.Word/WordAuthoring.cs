using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;

namespace Yadg.Word;

public sealed class WordAuthoringException(Diagnostic diagnostic) : Exception(diagnostic.ToString())
{
    public Diagnostic Diagnostic { get; } = diagnostic;
}

public static class WordAuthoring
{
    private static readonly Regex TagPattern = new(@"\{\{yadg:section:content:[^{}]+\}\}", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static IReadOnlyList<string> FindTags(string templatePath)
    {
        using var document = WordprocessingDocument.Open(templatePath, false);
        return document.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .SelectMany(p => TagPattern.Matches(LogicalText(p)).Select(m => m.Value))
            .ToArray();
    }

    public static void Author(string templatePath, string outputPath, YadgDocument model)
    {
        File.Copy(templatePath, outputPath, true);
        using var document = WordprocessingDocument.Open(outputPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            var textNodes = paragraph.Descendants<Text>().ToArray();
            var logical = string.Concat(textNodes.Select(t => t.Text));
            var matches = TagPattern.Matches(logical).Cast<Match>().ToArray();
            foreach (var match in matches.Reverse())
            {
                if (!SectionReference.TryParseTag(match.Value, out var reference, out var tagDiagnostic))
                    throw new WordAuthoringException(tagDiagnostic!);
                var resolved = MarkdownDocumentParser.Resolve(model, reference!);
                if (resolved.Diagnostic is not null) throw new WordAuthoringException(resolved.Diagnostic);
                ReplaceLogicalRange(textNodes, match.Index, match.Length, resolved.Content!);
                found.Add(match.Value);
            }
        }
        if (found.Count == 0)
            throw new WordAuthoringException(new("YADG-TAG-002", "Template contains no supported visible YADG section-content tag."));
        document.MainDocumentPart.Document.Save();
    }

    private static string LogicalText(Paragraph paragraph) => string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));

    private static void ReplaceLogicalRange(IReadOnlyList<Text> nodes, int start, int length, string replacement)
    {
        var end = start + length;
        var position = 0;
        var first = -1;
        for (var i = 0; i < nodes.Count; i++)
        {
            var nodeStart = position;
            var nodeEnd = position + nodes[i].Text.Length;
            if (first < 0 && start >= nodeStart && start < nodeEnd) first = i;
            if (first >= 0 && nodeStart < end && nodeEnd > start)
            {
                var localStart = Math.Max(start - nodeStart, 0);
                var localEnd = Math.Min(end - nodeStart, nodes[i].Text.Length);
                var before = nodes[i].Text[..localStart];
                var after = nodes[i].Text[localEnd..];
                nodes[i].Text = i == first ? before + replacement + after : before + after;
            }
            position = nodeEnd;
        }
        if (first < 0) throw new InvalidOperationException("Logical tag range did not map to OOXML text nodes.");
    }
}
