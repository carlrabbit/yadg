using System.Text.RegularExpressions;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Yadg.Core;

namespace Yadg.Presentation;

public static class PresentationMarkdownParser
{
    private static readonly Regex HeadingLine = new(@"^(?<marks>#{1,6})\s+(?<text>.*?)(?:\s+\{#(?<id>[^}]+)\})?\s*#*\s*$", RegexOptions.Compiled);
    private static readonly Regex ValidSelector = new("^[A-Za-z][A-Za-z0-9_-]*$", RegexOptions.Compiled);
    private static readonly Regex MermaidInfo = new("^mermaid\\s+\\{#(?<id>[A-Za-z][A-Za-z0-9_-]*)(?:\\s+caption=\"(?<caption>[^\"]*)\")?\\}$", RegexOptions.Compiled);

    public static PresentationDocument Parse(IEnumerable<PresentationSource> sources, string inlineCodePolicy = "style", string thematicBreakPolicy = "error", bool yolo = false)
    {
        var diagnostics = new List<Diagnostic>();
        var slides = new List<PresentationSlide>();
        var pipeline = new MarkdownPipelineBuilder().UsePipeTables().UseGenericAttributes().UseEmphasisExtras(EmphasisExtraOptions.Strikethrough).Build();
        foreach (var source in sources)
        {
            var effectiveCodePolicy = inlineCodePolicy;
            if (yolo && inlineCodePolicy == "error")
            {
                foreach (Match _ in Regex.Matches(source.Content, "(?<!`)`[^`\\r\\n]+`(?!`)"))
                    diagnostics.Add(new("YADG-YOLO-CODE-001", "Inline code is prohibited by markdown.codeInline: error; fallback: preserve the code text as plain text.", false, source.Path) { IsDegradation = true });
                effectiveCodePolicy = "ignore";
            }
            var sourceSlides = Markdown.Parse(source.Content, pipeline);
            var lines = source.Content.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            SlideBuilder? current = null;
            var currentRegion = (string?)null;
            var direct = new List<PresentationBlock>();
            var regions = new Dictionary<string, List<PresentationBlock>>(StringComparer.Ordinal);

            void FinishSlide()
            {
                if (current is null) return;
                if (!PresentationLayouts.All.TryGetValue(current.Layout, out var layout))
                {
                    diagnostics.Add(new("YADG-PPT-MD-002", $"Unknown presentation layout '{current.Layout}'.", true, Location(source.Path, current.Line)));
                    current = null;
                    return;
                }
                if (layout.Regions.Count == 0)
                {
                    if (direct.Count > 0 || regions.Count > 0) diagnostics.Add(new("YADG-PPT-MD-003", $"Slide '{current.Title}' uses layout 'section', which has no content regions.", true, Location(source.Path, current.Line)));
                }
                else if (layout.Regions.Count == 1)
                {
                    var onlyRegion = layout.Regions.Keys.Single();
                    if (direct.Count > 0 && regions.ContainsKey(onlyRegion))
                        diagnostics.Add(new("YADG-PPT-MD-004", $"Slide '{current.Title}' supplies both direct body content and region '{onlyRegion}'.", true, Location(source.Path, current.Line)));
                    else if (direct.Count > 0) regions[onlyRegion] = direct.ToList();
                    else if (!regions.ContainsKey(onlyRegion)) diagnostics.Add(new("YADG-PPT-MD-005", $"Slide '{current.Title}' is missing required region '{onlyRegion}'.", true, Location(source.Path, current.Line)));
                }
                else
                {
                    if (direct.Count > 0) diagnostics.Add(new("YADG-PPT-MD-006", $"Slide '{current.Title}' uses multiple regions; all body content must be under a level-2 region selector.", true, Location(source.Path, current.Line)));
                    foreach (var expected in layout.Regions.Keys)
                        if (!regions.ContainsKey(expected)) diagnostics.Add(new("YADG-PPT-MD-005", $"Slide '{current.Title}' is missing required region '{expected}'.", true, Location(source.Path, current.Line)));
                }
                foreach (var actual in regions.Keys)
                    if (!layout.Regions.ContainsKey(actual)) diagnostics.Add(new("YADG-PPT-MD-007", $"Slide '{current.Title}' selects undefined region '{actual}' for layout '{layout.Id}'.", true, Location(source.Path, current.Line)));

                var parsedRegions = new List<PresentationRegion>();
                foreach (var (regionId, regionType) in layout.Regions)
                {
                    if (!regions.TryGetValue(regionId, out var content)) continue;
                    ValidatePayload(content, regionType, current, regionId, source.Path, diagnostics);
                    parsedRegions.Add(new(regionId, regionType, content.ToArray()));
                }
                slides.Add(new PresentationSlide($"{Path.GetFileName(source.Path)}:{slides.Count + 1}", current.Title, layout.Id, source.Path, current.Ordinal, parsedRegions) { TitleInlines = current.TitleInlines });
                current = null;
            }

            foreach (var block in sourceSlides)
            {
                if (block is HeadingBlock heading)
                {
                    var headingLine = block.Line >= 0 && block.Line < lines.Length ? HeadingLine.Match(lines[block.Line].TrimEnd()) : Match.Empty;
                    var text = InlineText(heading.Inline);
                    var selector = headingLine.Success ? headingLine.Groups["id"].Value : string.Empty;
                    text = Regex.Replace(text, @"\s*\{#[^}]+\}\s*$", string.Empty).TrimEnd();
                    if (heading.Level == 1)
                    {
                        FinishSlide();
                        direct = new List<PresentationBlock>();
                        regions = new Dictionary<string, List<PresentationBlock>>(StringComparer.Ordinal);
                        currentRegion = null;
                        if (!ValidSelector.IsMatch(selector) && selector.Length > 0)
                            diagnostics.Add(new("YADG-PPT-MD-001", $"Invalid layout selector '{selector}'.", true, Location(source.Path, block.Line)));
                        var titleInlines = ConvertInlines(heading.Inline, effectiveCodePolicy, source.Path, diagnostics);
                        current = new SlideBuilder(text, selector.Length == 0 ? "content" : selector, block.Line + 1, slides.Count + 1, titleInlines);
                        continue;
                    }
                    if (current is null)
                    {
                        diagnostics.Add(new("YADG-PPT-MD-008", "Each presentation source must begin with a level-1 slide heading; stray content before the first slide is unsupported.", true, Location(source.Path, block.Line)));
                        continue;
                    }
                    if (heading.Level == 2)
                    {
                        if (selector.Length == 0 || !ValidSelector.IsMatch(selector))
                        {
                            diagnostics.Add(new("YADG-PPT-MD-009", "Every level-2 region heading requires a valid {#region-id} selector.", true, Location(source.Path, block.Line)));
                            currentRegion = null;
                            continue;
                        }
                        if (!PresentationLayouts.All.TryGetValue(current.Layout, out var selectedLayout) || !selectedLayout.Regions.ContainsKey(selector))
                            diagnostics.Add(new("YADG-PPT-MD-007", $"Slide '{current.Title}' selects undefined region '{selector}' for layout '{current.Layout}'.", true, Location(source.Path, block.Line)));
                        if (!regions.TryAdd(selector, new List<PresentationBlock>()))
                            diagnostics.Add(new("YADG-PPT-MD-010", $"Slide '{current.Title}' selects region '{selector}' more than once.", true, Location(source.Path, block.Line)));
                        currentRegion = selector;
                        continue;
                    }
                    AddBlock(new PresentationParagraph(ConvertInlines(heading.Inline, effectiveCodePolicy, source.Path, diagnostics), heading.Level));
                    continue;
                }
                if (current is null)
                {
                    if (!string.IsNullOrWhiteSpace(block.ToString())) diagnostics.Add(new("YADG-PPT-MD-008", "Each presentation source must begin with a level-1 slide heading; stray content before the first slide is unsupported.", true, Location(source.Path, block.Line)));
                    continue;
                }
                if (block is ThematicBreakBlock)
                {
                    if (thematicBreakPolicy == "error")
                    {
                        if (yolo) diagnostics.Add(new("YADG-YOLO-MD-001", "Markdown thematic break omitted; fallback: no thematic-break output.", false, Location(source.Path, block.Line)) { IsDegradation = true });
                        else diagnostics.Add(new("YADG-PPT-MD-019", "Unsupported Markdown construct 'thematic break'.", true, Location(source.Path, block.Line)));
                    }
                    continue;
                }
                var converted = ConvertBlock(block, effectiveCodePolicy, source.Path, diagnostics, lines);
                if (converted is not null) AddBlock(converted);
            }
            FinishSlide();

            void AddBlock(PresentationBlock content)
            {
                if (currentRegion is null) direct.Add(content);
                else if (regions.TryGetValue(currentRegion, out var regionBlocks)) regionBlocks.Add(content);
            }
        }
        var generatedIds = slides.SelectMany(slide => slide.Regions).SelectMany(region => region.Blocks).OfType<PresentationImage>().Where(image => image.Id is not null).GroupBy(image => image.Id!, StringComparer.Ordinal);
        foreach (var duplicate in generatedIds.Where(group => group.Count() > 1)) diagnostics.Add(new("YADG-PPT-MERMAID-001", $"Mermaid ID '{duplicate.Key}' must be unique within the presentation deck.", true, duplicate.Skip(1).First().SourcePath));
        if (slides.Count == 0 && diagnostics.All(d => !d.IsError)) diagnostics.Add(new("YADG-PPT-MD-011", "Presentation has no level-1 slide headings."));
        return new(slides, diagnostics);
    }

    private static PresentationBlock? ConvertBlock(Block block, string inlineCodePolicy, string sourcePath, List<Diagnostic> diagnostics, string[]? sourceLines = null)
    {
        if (block is FencedCodeBlock fenced)
        {
            var info = fenced.Info?.Trim() ?? string.Empty;
            if (sourceLines is not null && block.Line >= 0 && block.Line < sourceLines.Length)
            {
                var opening = Regex.Match(sourceLines[block.Line], @"^\s*(?:`{3,}|~{3,})(?<info>.*)$");
                if (opening.Success) info = opening.Groups["info"].Value.Trim();
            }
            if (info.StartsWith("mermaid", StringComparison.OrdinalIgnoreCase))
            {
                var match = MermaidInfo.Match(info);
                if (!match.Success) { diagnostics.Add(new("YADG-PPT-MERMAID-002", "Mermaid fence requires exactly a stable ID and optional caption attribute.", true, Location(sourcePath, block.Line))); return null; }
                var code = fenced.Lines.ToString().Trim();
                if (code.Length == 0) { diagnostics.Add(new("YADG-PPT-MERMAID-003", "Mermaid source must not be empty.", true, Location(sourcePath, block.Line))); return null; }
                return new PresentationImage(match.Groups["caption"].Value, string.Empty, sourcePath, code, match.Groups["id"].Value, match.Groups["caption"].Success ? match.Groups["caption"].Value : null);
            }
        }
        if (block is ParagraphBlock paragraph)
        {
            var images = paragraph.Inline?.OfType<LinkInline>().Where(link => link.IsImage).ToArray() ?? Array.Empty<LinkInline>();
            if (images.Length > 0)
            {
                if (images.Length != 1 || paragraph.Inline!.Count() != 1)
                {
                    diagnostics.Add(new("YADG-PPT-MD-012", "An image region must contain one image block and no other content.", true, Location(sourcePath, block.Line)));
                    return null;
                }
                return new PresentationImage(InlineText(images[0]), images[0].Url ?? string.Empty, sourcePath);
            }
            return new PresentationParagraph(ConvertInlines(paragraph.Inline, inlineCodePolicy, sourcePath, diagnostics));
        }
        if (block is HeadingBlock) return null;
        if (block is ListBlock list) return ConvertList(list, inlineCodePolicy, sourcePath, diagnostics);
        if (block is Table table)
        {
            var rows = table.OfType<TableRow>().ToArray();
            if (rows.Length < 2)
            {
                diagnostics.Add(new("YADG-PPT-MD-013", "A presentation table requires a header row and at least one body row.", true, Location(sourcePath, block.Line)));
                return null;
            }
            var header = rows[0].OfType<TableCell>().Select(cell => CellInlines(cell, inlineCodePolicy, sourcePath, diagnostics)).ToArray();
            var body = rows.Skip(1).Select(row => (IReadOnlyList<IReadOnlyList<PresentationInline>>)row.OfType<TableCell>().Select(cell => CellInlines(cell, inlineCodePolicy, sourcePath, diagnostics)).ToArray()).ToArray();
            return new PresentationTable(header, body);
        }
        diagnostics.Add(new("YADG-PPT-MD-014", $"Unsupported presentation block '{block.GetType().Name}'.", true, Location(sourcePath, block.Line)));
        return null;
    }

    private static PresentationList ConvertList(ListBlock list, string inlineCodePolicy, string sourcePath, List<Diagnostic> diagnostics)
    {
        var items = new List<PresentationListItem>();
        foreach (var item in list.OfType<ListItemBlock>())
        {
            var blocks = new List<PresentationBlock>();
            foreach (var child in item)
            {
                if (child is ParagraphBlock paragraph) blocks.Add(new PresentationParagraph(ConvertInlines(paragraph.Inline, inlineCodePolicy, sourcePath, diagnostics)));
                else if (child is ListBlock nested) blocks.Add(ConvertList(nested, inlineCodePolicy, sourcePath, diagnostics));
                else diagnostics.Add(new("YADG-PPT-MD-014", $"Unsupported list item block '{child.GetType().Name}'.", true, Location(sourcePath, child.Line)));
            }
            items.Add(new(blocks));
        }
        return new(list.IsOrdered, items);
    }

    private static IReadOnlyList<PresentationInline> CellInlines(TableCell cell, string inlineCodePolicy, string sourcePath, List<Diagnostic> diagnostics) =>
        cell.OfType<ParagraphBlock>().SelectMany(paragraph => ConvertInlines(paragraph.Inline, inlineCodePolicy, sourcePath, diagnostics)).ToArray();

    private static IReadOnlyList<PresentationInline> ConvertInlines(ContainerInline? container, string inlineCodePolicy, string sourcePath, List<Diagnostic> diagnostics)
    {
        var result = new List<PresentationInline>();
        if (container is null) return result;
        foreach (var inline in container)
        {
            switch (inline)
            {
                case LiteralInline literal: result.Add(new PresentationText(literal.Content.ToString())); break;
                case LineBreakInline lineBreak: result.Add(new PresentationBreak(lineBreak.IsHard)); break;
                case EmphasisInline emphasis when emphasis.DelimiterChar == '~': result.Add(new PresentationStrike(ConvertInlines(emphasis, inlineCodePolicy, sourcePath, diagnostics))); break;
                case EmphasisInline emphasis: result.Add(new PresentationEmphasis(ConvertInlines(emphasis, inlineCodePolicy, sourcePath, diagnostics), emphasis.DelimiterCount >= 2, emphasis.DelimiterCount % 2 == 1)); break;
                case LinkInline link when !link.IsImage:
                    var destination = link.Url ?? string.Empty;
                    if (!Uri.TryCreate(destination, UriKind.Absolute, out var uri) || uri.Scheme is not "http" and not "https" and not "mailto")
                        diagnostics.Add(new("YADG-PPT-LINK-001", $"Link destination '{destination}' is unsupported; visible text is retained without a hyperlink.", false, sourcePath));
                    result.Add(new PresentationLink(destination, ConvertInlines(link, inlineCodePolicy, sourcePath, diagnostics)));
                    break;
                case CodeInline code when inlineCodePolicy == "error":
                    if (diagnostics.All(d => d.Code != "YADG-PPT-MD-020" || d.Location != sourcePath)) diagnostics.Add(new("YADG-PPT-MD-020", "Inline code is prohibited by markdown.codeInline: error.", true, sourcePath));
                    result.Add(new PresentationText(code.Content));
                    break;
                case CodeInline code when inlineCodePolicy == "ignore": result.Add(new PresentationText(code.Content)); break;
                case CodeInline code: result.Add(new PresentationCode(code.Content)); break;
                default: diagnostics.Add(new("YADG-PPT-MD-015", $"Unsupported presentation inline '{inline.GetType().Name}'.", true, sourcePath)); break;
            }
        }
        return result;
    }

    private static void ValidatePayload(IReadOnlyList<PresentationBlock> blocks, PresentationRegionType type, SlideBuilder slide, string region, string source, List<Diagnostic> diagnostics)
    {
        if (blocks.Count == 0) { diagnostics.Add(new("YADG-PPT-MD-016", $"Slide '{slide.Title}' region '{region}' requires content.", true, Location(source, slide.Line))); return; }
        var valid = type switch
        {
            PresentationRegionType.Text => blocks.All(block => block is PresentationParagraph or PresentationList),
            PresentationRegionType.Image => blocks.Count == 1 && blocks[0] is PresentationImage image && (image.GeneratedSource is not null && image.Id is not null || IsLocalImage(image.Path)),
            PresentationRegionType.Table => blocks.Count == 1 && blocks[0] is PresentationTable,
            _ => false
        };
        if (!valid)
        {
            var description = type switch
            {
                PresentationRegionType.Text => "accepts paragraphs, content headings and lists only",
                PresentationRegionType.Image => "requires exactly one local PNG/JPEG image or one identified Mermaid block",
                _ => "requires exactly one pipe table with header and body rows"
            };
            diagnostics.Add(new("YADG-PPT-MD-017", $"Slide '{slide.Title}' region '{region}' {description}.", true, Location(source, slide.Line)));
        }
    }

    private static bool IsLocalImage(string path) => !string.IsNullOrWhiteSpace(path) && !Uri.TryCreate(path, UriKind.Absolute, out _) && !path.Contains('*') && !path.Contains('?') && Path.GetExtension(path).ToLowerInvariant() is ".png" or ".jpg" or ".jpeg";
    private static string InlineText(ContainerInline? inline) => inline is null ? string.Empty : string.Concat(inline.Select(child => child switch { LiteralInline literal => literal.Content.ToString(), CodeInline code => code.Content, ContainerInline nested => InlineText(nested), _ => string.Empty }));
    private static string Location(string source, int line) => $"{source}:{Math.Max(1, line + 1)}";
    private sealed record SlideBuilder(string Title, string Layout, int Line, int Ordinal, IReadOnlyList<PresentationInline> TitleInlines);
}
