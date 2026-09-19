using System.Text.RegularExpressions;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Yadg.Core;

public static class MarkdownDocumentParser
{
    private static readonly Regex StableId = new(@"^(?<text>.*?)\s*\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\s*$", RegexOptions.Compiled);

    public static ParseResult Parse(string markdown, string location = "<input>")
    {
        var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
        var ast = Markdown.Parse(markdown, pipeline);
        var diagnostics = new List<Diagnostic>();
        var blocks = new List<YadgBlock>();
        var headingMetadata = markdown.Replace("\r\n", "\n").Split('\n')
            .Select(line => Regex.Match(line, @"^(?<marks>#{1,6})\s+(?<text>.+?)(?:\s*\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)\})?\s*$"))
            .Where(match => match.Success)
            .Select(match => (Text: Regex.Replace(match.Groups["text"].Value.Trim(), @"\s*\{#[A-Za-z][A-Za-z0-9_-]*\}$", ""), Id: match.Groups["id"].Success ? match.Groups["id"].Value : null))
            .ToArray();
        var headingOrdinal = 0;
        foreach (var block in ast)
        {
            var converted = ConvertBlock(block, diagnostics, location, headingMetadata, ref headingOrdinal);
            if (converted is not null) blocks.Add(converted);
        }
        var sections = new Dictionary<string, YadgSection>(StringComparer.Ordinal);
        for (var i = 0; i < blocks.Count; i++)
        {
            if (blocks[i] is not YadgHeading { Id: not null } heading) continue;
            var end = blocks.Count;
            for (var j = i + 1; j < blocks.Count; j++)
                if (blocks[j] is YadgHeading next && next.Level <= heading.Level) { end = j; break; }
            var section = new YadgSection(heading.Id!, heading, blocks[(i + 1)..end]);
            if (!sections.TryAdd(section.Id, section)) diagnostics.Add(new("YADG-REF-002", $"Duplicate stable section ID '{section.Id}'.", true, location));
        }
        return new(new YadgDocument(blocks, sections), diagnostics);
    }

    public static YadgDocument Merge(IEnumerable<(string Location, ParseResult Result)> sources, List<Diagnostic> diagnostics)
    {
        var allBlocks = new List<YadgBlock>();
        var registry = new Dictionary<string, YadgSection>(StringComparer.Ordinal);
        foreach (var (location, result) in sources.OrderBy(s => s.Location, StringComparer.Ordinal))
        {
            diagnostics.AddRange(result.Diagnostics);
            if (result.Document is null) continue;
            allBlocks.AddRange(result.Document.Blocks);
            foreach (var section in result.Document.References.Values)
                if (!registry.TryAdd(section.Id, section)) diagnostics.Add(new("YADG-REF-002", $"Duplicate stable section ID '{section.Id}' across workspace sources.", true, location));
        }
        return new(allBlocks, registry);
    }

    private static YadgBlock? ConvertBlock(Block block, List<Diagnostic> diagnostics, string location, IReadOnlyList<(string Text, string? Id)> headingMetadata, ref int headingOrdinal)
    {
        if (block is LinkReferenceDefinitionGroup) return null;
        if (block is HeadingBlock heading)
        {
            var text = InlineText(heading.Inline);
            var match = StableId.Match(text);
            var hasMetadata = headingOrdinal < headingMetadata.Count;
            var metadata = hasMetadata ? headingMetadata[headingOrdinal] : (Text: (string?)null, Id: (string?)null);
            headingOrdinal++;
            return hasMetadata ? new YadgHeading(heading.Level, metadata.Text!, metadata.Id) : match.Success ? new YadgHeading(heading.Level, match.Groups["text"].Value, match.Groups["id"].Value) : new YadgHeading(heading.Level, text);
        }
        if (block is ParagraphBlock paragraph)
            return new YadgParagraph(ConvertInlines(paragraph.Inline, diagnostics, location));
        diagnostics.Add(new("YADG-MD-UNSUPPORTED", $"Unsupported Markdown construct '{block.GetType().Name}'.", true, location));
        return null;
    }

    private static IReadOnlyList<YadgInline> ConvertInlines(ContainerInline? container, List<Diagnostic> diagnostics, string location)
    {
        var result = new List<YadgInline>();
        if (container is null) return result;
        foreach (var inline in container)
        {
            switch (inline)
            {
                case LiteralInline literal:
                    var value = literal.Content.ToString();
                    if (value.Contains('\n'))
                    {
                        var parts = value.Split('\n');
                        for (var i = 0; i < parts.Length; i++) { if (parts[i].Length > 0) result.Add(new YadgText(parts[i])); if (i < parts.Length - 1) result.Add(new YadgSoftBreak()); }
                    }
                    else result.Add(new YadgText(value));
                    break;
                case LineBreakInline lineBreak:
                    result.Add(lineBreak.IsHard ? new YadgHardBreak() : new YadgSoftBreak());
                    break;
                case EmphasisInline emphasis:
                    var children = ConvertInlines(emphasis, diagnostics, location);
                    result.Add(emphasis.DelimiterCount >= 2 ? new YadgStrong(children) : new YadgEmphasis(children));
                    break;
                default:
                    diagnostics.Add(new("YADG-MD-UNSUPPORTED", $"Unsupported Markdown inline '{inline.GetType().Name}'.", true, location));
                    break;
            }
        }
        return result;
    }

    private static string InlineText(ContainerInline? inline) => inline is null ? string.Empty : string.Concat(inline.Select(i => i switch
    {
        LiteralInline literal => literal.Content.ToString(),
        EmphasisInline emphasis => InlineText(emphasis),
        LineBreakInline => " ",
        _ => string.Empty
    }));

    public static (IReadOnlyList<YadgBlock>? Blocks, Diagnostic? Diagnostic) Resolve(YadgDocument document, SectionReference reference, string location = "<template>")
    {
        var section = document.FindSection(reference.Id);
        return section is null ? (null, new("YADG-REF-001", $"Unresolved stable section ID '{reference.Id}'.", true, location)) : (section.Select(reference.Selection), null);
    }
}
