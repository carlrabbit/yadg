using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Markdig;
using Markdig.Extensions.Tables;
using Yadg.Core;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Yadg.Presentation;

public sealed record PresentationAnalysis(PresentationDocument Document, IReadOnlyList<Diagnostic> Diagnostics) : IDisposable
{
    public IReadOnlyList<string> TemporaryProducerDirectories { get; init; } = Array.Empty<string>();
    public IReadOnlyDictionary<string, string> Values { get; init; } = new Dictionary<string, string>(StringComparer.Ordinal);
    public bool Yolo { get; init; }
    public bool IsValid => Diagnostics.All(diagnostic => !diagnostic.IsError) && Document.IsValid;
    public void Dispose()
    {
        foreach (var directory in TemporaryProducerDirectories.Distinct(StringComparer.OrdinalIgnoreCase))
            try { if (Directory.Exists(directory)) Directory.Delete(directory, true); } catch { }
    }
}

public static class PresentationAuthoring
{
    private static readonly Regex SlideMarker = new(@"^\{\{slide:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$", RegexOptions.Compiled);
    private static readonly Regex RegionMarker = new(@"^\{\{(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$", RegexOptions.Compiled);
    private static readonly Regex AnyControl = new(@"\{\{[^{}]*\}\}", RegexOptions.Compiled);

    public static PresentationAnalysis Analyze(string templatePath, IEnumerable<PresentationSource> sources, string effectiveRoot, string inlineCodePolicy = "style", string thematicBreakPolicy = "error", IReadOnlyDictionary<string, string>? values = null, bool yolo = false, IReadOnlyDictionary<string, MermaidProducerConfiguration>? producers = null)
    {
        var document = PresentationMarkdownParser.Parse(sources, inlineCodePolicy, thematicBreakPolicy, yolo);
        var diagnostics = document.Diagnostics.ToList();
        if (values is not null) document = SubstituteSourceValues(document, values, yolo, diagnostics);
        if (!File.Exists(templatePath) || !Path.GetExtension(templatePath).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
        {
            diagnostics.Add(new("YADG-PPT-TEMPLATE-001", "Presentation template must be an existing .pptx file.", true, templatePath));
            return new(document, diagnostics);
        }
        Dictionary<string, SlidePart> prototypes;
        try
        {
            using var presentation = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(templatePath, false);
            prototypes = DiscoverPrototypes(presentation.PresentationPart!, diagnostics, templatePath);
            ValidateSelectedPrototypes(presentation.PresentationPart!, document, prototypes, diagnostics, templatePath, yolo);
            foreach (var slide in document.Slides)
                if (prototypes.TryGetValue(slide.Layout, out var prototype)) ValidatePrototypeValues(prototype, values ?? new Dictionary<string, string>(StringComparer.Ordinal), slide.SourcePath, yolo, diagnostics);
        }
        catch (Exception ex) { diagnostics.Add(new("YADG-PPT-TEMPLATE-002", $"Unable to inspect PowerPoint template: {ex.Message}", true, templatePath)); return new(document, diagnostics); }

        var temporary = new List<string>();
        try
        {
            document = ResolveImages(document, effectiveRoot, producers, diagnostics, temporary, yolo);
            return new(document, diagnostics) { TemporaryProducerDirectories = temporary, Values = values ?? new Dictionary<string, string>(StringComparer.Ordinal), Yolo = yolo };
        }
        catch
        {
            foreach (var directory in temporary) try { if (Directory.Exists(directory)) Directory.Delete(directory, true); } catch { }
            throw;
        }
    }

    private static PresentationDocument SubstituteSourceValues(PresentationDocument document, IReadOnlyDictionary<string, string> values, bool yolo, List<Diagnostic> diagnostics)
    {
        PresentationInline Inline(PresentationInline inline, string location) => inline switch
        {
            PresentationText text => text with { Value = ValueSubstitution.Resolve(text.Value, values, location, yolo, diagnostics) },
            PresentationCode code => code with { Value = ValueSubstitution.Resolve(code.Value, values, location, yolo, diagnostics) },
            PresentationEmphasis emphasis => emphasis with { Children = emphasis.Children.Select(child => Inline(child, location)).ToArray() },
            PresentationStrike strike => strike with { Children = strike.Children.Select(child => Inline(child, location)).ToArray() },
            PresentationLink link => link with { Children = link.Children.Select(child => Inline(child, location)).ToArray() },
            _ => inline
        };
        PresentationBlock Block(PresentationBlock block, string location) => block switch
        {
            PresentationParagraph paragraph => paragraph with { Inlines = paragraph.Inlines.Select(inline => Inline(inline, location)).ToArray() },
            PresentationList list => list with { Items = list.Items.Select(item => item with { Blocks = item.Blocks.Select(child => Block(child, location)).ToArray() }).ToArray() },
            PresentationTable table => table with { Header = table.Header.Select(cell => (IReadOnlyList<PresentationInline>)cell.Select(inline => Inline(inline, location)).ToArray()).ToArray(), Rows = table.Rows.Select(row => (IReadOnlyList<IReadOnlyList<PresentationInline>>)row.Select(cell => (IReadOnlyList<PresentationInline>)cell.Select(inline => Inline(inline, location)).ToArray()).ToArray()).ToArray() },
            _ => block
        };
        string Plain(PresentationInline inline) => inline switch
        {
            PresentationText text => text.Value,
            PresentationCode code => code.Value,
            PresentationBreak => " ",
            PresentationEmphasis emphasis => string.Concat(emphasis.Children.Select(Plain)),
            PresentationStrike strike => string.Concat(strike.Children.Select(Plain)),
            PresentationLink link => string.Concat(link.Children.Select(Plain)),
            _ => string.Empty
        };
        var slides = document.Slides.Select(slide =>
        {
            var titleLocation = $"{slide.SourcePath}: slide '{slide.Title}' title";
            var titleInlines = slide.TitleInlines.Select(inline => Inline(inline, titleLocation)).ToArray();
            return slide with
            {
                TitleInlines = titleInlines,
                Title = string.Concat(titleInlines.Select(Plain)),
                Regions = slide.Regions.Select(region => region with { Blocks = region.Blocks.Select(block => Block(block, $"{slide.SourcePath}: slide '{slide.Title}' region '{region.Id}'")).ToArray() }).ToArray()
            };
        }).ToArray();
        return document with { Slides = slides, Diagnostics = diagnostics };
    }

    public static IReadOnlyList<Diagnostic> Build(string templatePath, IEnumerable<PresentationSource> sources, string effectiveRoot, string outputPath, string inlineCodePolicy = "style", string thematicBreakPolicy = "error", IReadOnlyDictionary<string, string>? values = null, bool yolo = false, IReadOnlyDictionary<string, MermaidProducerConfiguration>? producers = null)
    {
        var sourceArray = sources.ToArray();
        using var analysis = Analyze(templatePath, sourceArray, effectiveRoot, inlineCodePolicy, thematicBreakPolicy, values, yolo, producers);
        return BuildAnalyzed(templatePath, analysis, effectiveRoot, outputPath);
    }

    public static IReadOnlyList<Diagnostic> BuildAnalyzed(string templatePath, PresentationAnalysis analysis, string effectiveRoot, string outputPath)
    {
        if (!analysis.IsValid) return analysis.Diagnostics;
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
        File.Copy(templatePath, outputPath, true);
        using var package = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(outputPath, true);
        var presentationPart = package.PresentationPart!;
        var prototypes = DiscoverPrototypes(presentationPart, new List<Diagnostic>(), templatePath);
        var generated = new List<(SlidePart Part, string RelationshipId)>();
        foreach (var slide in analysis.Document.Slides)
        {
            var sourcePrototype = prototypes[slide.Layout];
            var clone = CloneSlide(presentationPart, sourcePrototype);
            var diagnostics = new List<Diagnostic>();
            SubstitutePrototypeValues(clone, analysis.Values, slide.SourcePath, analysis.Yolo, diagnostics);
            AuthorSlide(clone, slide, effectiveRoot, diagnostics);
            if (diagnostics.Any(diagnostic => diagnostic.IsError)) throw new InvalidDataException(string.Join(Environment.NewLine, diagnostics));
            var relationshipId = presentationPart.GetIdOfPart(clone);
            generated.Add((clone, relationshipId));
        }
        var generatedParts = generated.Select(item => item.Part).ToHashSet();
        foreach (var unselected in presentationPart.SlideParts.Where(part => !generatedParts.Contains(part)).ToArray())
            presentationPart.DeletePart(unselected);
        var slideList = presentationPart.Presentation.SlideIdList ?? presentationPart.Presentation.AppendChild(new SlideIdList());
        slideList.RemoveAllChildren<SlideId>();
        uint id = 256;
        foreach (var item in generated) slideList.AppendChild(new SlideId { Id = id++, RelationshipId = item.RelationshipId });
        presentationPart.Presentation.Save();
        return analysis.Diagnostics;
    }

    private static PresentationDocument ResolveImages(PresentationDocument document, string root, IReadOnlyDictionary<string, MermaidProducerConfiguration>? producers, List<Diagnostic> diagnostics, List<string> temporary, bool yolo)
    {
        var slides = document.Slides.Select(slide => slide with
        {
            Regions = slide.Regions.Select(region => region with
            {
                Blocks = region.Blocks.Select(block =>
                {
                    if (block is not PresentationImage image) return block;
                    if (image.GeneratedSource is not null)
                    {
                        var before = diagnostics.Count;
                        var config = producers is not null && producers.TryGetValue("mermaid", out var producer) ? producer : null;
                        var asset = MermaidProducer.RenderSource(image.Id ?? string.Empty, image.GeneratedSource, image.SourcePath, config, root, diagnostics, temporary);
                        if (asset is not null) return image with { Path = asset.FullPath };
                        if (!yolo) return image;
                        for (var i = before; i < diagnostics.Count; i++) diagnostics[i] = diagnostics[i] with { Code = "YADG-YOLO-PRODUCER-001", IsError = false, IsDegradation = true, Message = diagnostics[i].Message + "; fallback: visible image-region placeholder." };
                        return image with { FailureReason = "Mermaid producer unavailable or failed" };
                    }
                    var path = ResolveImage(root, image, out var error);
                    if (path is not null) return image;
                    if (yolo && error is not null && error.Contains("asset does not exist", StringComparison.OrdinalIgnoreCase))
                    {
                        diagnostics.Add(new("YADG-YOLO-IMAGE-001", $"Image asset is missing; fallback: visible image-region placeholder. {error}", false, image.SourcePath) { IsDegradation = true });
                        return image with { FailureReason = "image asset missing" };
                    }
                    diagnostics.Add(new("YADG-PPT-ASSET-001", error!, true, image.SourcePath));
                    return image;
                }).ToArray()
            }).ToArray()
        }).ToArray();
        return document with { Slides = slides, Diagnostics = diagnostics };
    }

    private static readonly Regex ValueToken = new(@"\{\{value:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}", RegexOptions.Compiled);

    private static void SubstitutePrototypeValues(SlidePart slide, IReadOnlyDictionary<string, string> values, string sourcePath, bool yolo, List<Diagnostic> diagnostics)
    {
        foreach (var paragraph in slide.Slide.Descendants<A.Paragraph>())
        {
            var textNodes = paragraph.Descendants<A.Run>().SelectMany(run => run.Elements<A.Text>()).ToArray();
            if (textNodes.Length == 0) continue;
            var original = string.Concat(textNodes.Select(node => node.Text));
            if (!original.Contains("{{value:", StringComparison.Ordinal)) continue;
            _ = ValueSubstitution.Resolve(original, values, sourcePath, yolo, diagnostics);
            var matches = ValueToken.Matches(original).Cast<Match>().Where(match => values.ContainsKey(match.Groups["id"].Value)).ToArray();
            if (matches.Length == 0) continue;
            var starts = new int[textNodes.Length];
            for (var i = 1; i < textNodes.Length; i++) starts[i] = starts[i - 1] + textNodes[i - 1].Text.Length;
            foreach (var match in matches.Reverse())
            {
                var startIndex = Array.FindLastIndex(starts, start => start <= match.Index);
                var endOffset = match.Index + match.Length;
                var endIndex = Array.FindLastIndex(starts, start => start < endOffset);
                if (startIndex < 0 || endIndex < startIndex) continue;
                var firstOffset = match.Index - starts[startIndex];
                var lastEnd = endOffset - starts[endIndex];
                var replacement = values[match.Groups["id"].Value];
                var first = textNodes[startIndex];
                if (startIndex == endIndex)
                    first.Text = first.Text[..firstOffset] + replacement + first.Text[lastEnd..];
                else
                {
                    first.Text = first.Text[..firstOffset] + replacement;
                    for (var i = startIndex + 1; i < endIndex; i++) textNodes[i].Text = string.Empty;
                    textNodes[endIndex].Text = textNodes[endIndex].Text[lastEnd..];
                }
            }
        }
    }

    private static void ValidatePrototypeValues(SlidePart slide, IReadOnlyDictionary<string, string> values, string sourcePath, bool yolo, List<Diagnostic> diagnostics)
    {
        foreach (var paragraph in slide.Slide.Descendants<A.Paragraph>())
        {
            var text = string.Concat(paragraph.Descendants<A.Run>().SelectMany(run => run.Elements<A.Text>()).Select(node => node.Text));
            if (text.Contains("{{value:", StringComparison.Ordinal)) _ = ValueSubstitution.Resolve(text, values, sourcePath, yolo, diagnostics);
        }
    }

    private static Dictionary<string, SlidePart> DiscoverPrototypes(PresentationPart part, List<Diagnostic> diagnostics, string templatePath)
    {
        var result = new Dictionary<string, SlidePart>(StringComparer.Ordinal);
        foreach (var slidePart in part.SlideParts)
        {
            var markerParagraphs = slidePart.Slide.Descendants<A.Paragraph>().Select(paragraph => paragraph.InnerText).Where(text => text.Contains("{{", StringComparison.Ordinal) || text.Contains("}}", StringComparison.Ordinal)).ToArray();
            var markers = markerParagraphs;
            var slideMarkers = markers.Select(text => SlideMarker.Match(text)).Where(match => match.Success).ToArray();
            if (slideMarkers.Length == 0)
            {
                if (markers.Any(text => AnyControl.IsMatch(text))) diagnostics.Add(new("YADG-PPT-TEMPLATE-003", $"Malformed or unknown presentation control marker in template slide: {string.Join(", ", markers)}.", true, templatePath));
                continue;
            }
            var slideMarkerShapeIsDedicated = slidePart.Slide.Descendants<P.Shape>().Any(shape => SlideMarker.IsMatch(shape.TextBody?.InnerText ?? string.Empty));
            if (slideMarkers.Length != 1 || !slideMarkerShapeIsDedicated || markers.Any(text => AnyControl.IsMatch(text) && !SlideMarker.IsMatch(text) && text != "{{title}}" && !RegionMarker.IsMatch(text)))
            {
                diagnostics.Add(new("YADG-PPT-TEMPLATE-004", "Each prototype slide must contain exactly one dedicated {{slide:<layout-id>}} marker and only plain title/region markers.", true, templatePath));
                continue;
            }
            var layoutId = slideMarkers[0].Groups["id"].Value;
            if (!result.TryAdd(layoutId, slidePart)) diagnostics.Add(new("YADG-PPT-TEMPLATE-005", $"More than one prototype slide is marked for layout '{layoutId}'.", true, templatePath));
        }
        return result;
    }

    private static void ValidateSelectedPrototypes(PresentationPart package, PresentationDocument document, IReadOnlyDictionary<string, SlidePart> prototypes, List<Diagnostic> diagnostics, string templatePath, bool yolo = false)
    {
        foreach (var slide in document.Slides)
        {
            if (!prototypes.TryGetValue(slide.Layout, out var prototype))
            {
                diagnostics.Add(new("YADG-PPT-TEMPLATE-006", $"No prototype is marked for selected layout '{slide.Layout}'.", true, templatePath));
                continue;
            }
            var texts = prototype.Slide.Descendants<A.Text>().Select(text => text.Text).ToArray();
            var titleCount = texts.Count(text => text == "{{title}}");
            if (titleCount != 1) diagnostics.Add(new("YADG-PPT-TEMPLATE-007", $"Layout '{slide.Layout}' must contain exactly one plain {{title}} marker in a title shape.", true, templatePath));
            var titleShape = prototype.Slide.Descendants<P.Shape>().SingleOrDefault(shape => shape.TextBody?.InnerText.Contains("{{title}}", StringComparison.Ordinal) == true);
            if (titleShape is not null && RequiresCode(slide.TitleInlines) && CountCodeSamples(titleShape) != 1)
                AddCodeExemplarDiagnostic(diagnostics, "slide title", 0, CountCode(slide.TitleInlines), yolo, templatePath);
            foreach (var region in slide.Regions)
            {
                var token = $"{{{{{region.Id}}}}}";
                var count = texts.Count(text => text == token);
                if (count != 1) diagnostics.Add(new("YADG-PPT-TEMPLATE-008", $"Layout '{slide.Layout}' requires exactly one plain {token} marker.", true, templatePath));
                if (region.Type == PresentationRegionType.Image && !prototype.Slide.Descendants<P.Shape>().Any(shape => shape.TextBody?.InnerText.Contains(token, StringComparison.Ordinal) == true && shape.ShapeProperties?.Transform2D is not null))
                    diagnostics.Add(new("YADG-PPT-TEMPLATE-009", $"Image region '{region.Id}' must be a bounded text shape with a transform.", true, templatePath));
                if (region.Type == PresentationRegionType.Table)
                {
                    var frame = prototype.Slide.Descendants<P.GraphicFrame>().SingleOrDefault(candidate => candidate.Graphic?.GraphicData?.GetFirstChild<A.Table>()?.InnerText.Contains(token, StringComparison.Ordinal) == true);
                    var table = frame?.Graphic?.GraphicData?.GetFirstChild<A.Table>();
                    if (table is null)
                    {
                        diagnostics.Add(new("YADG-PPT-TEMPLATE-010", $"Table region '{region.Id}' must use a native editable table with its marker in the top-left cell.", true, templatePath));
                    }
                    else if (region.Blocks.OfType<PresentationTable>().SingleOrDefault() is { } content)
                    {
                        var rows = table.Elements<A.TableRow>().ToArray();
                        if (rows.Length != 2)
                            diagnostics.Add(new("YADG-PPT-TEMPLATE-015", "Table prototype must contain exactly one header example row and one body example row.", true, templatePath));
                        else if (rows[0].Elements<A.TableCell>().Count() != content.Header.Count || rows[1].Elements<A.TableCell>().Count() != content.Header.Count || content.Rows.Any(row => row.Count != content.Header.Count))
                            diagnostics.Add(new("YADG-PPT-TABLE-001", "Source table column count must match the native prototype table.", true, templatePath));
                        else if (table.Descendants<A.TableCell>().Any(cell => (cell.RowSpan?.Value ?? 1) > 1 || (cell.GridSpan?.Value ?? 1) > 1))
                            diagnostics.Add(new("YADG-PPT-TEMPLATE-016", "Merged prototype table cells are unsupported.", true, templatePath));
                        else
                        {
                            var shapeExtent = frame!.Transform?.Extents?.Cy?.Value;
                            var headerHeight = rows[0].Height?.Value ?? 0;
                            var bodyHeight = rows[1].Height?.Value ?? 0;
                            if (shapeExtent is > 0 && headerHeight > 0 && bodyHeight > 0 && headerHeight + bodyHeight * content.Rows.Count > shapeExtent.Value)
                                diagnostics.Add(new("YADG-PPT-TABLE-002", "Table rows cannot fit within the prototype's fixed structural height while preserving example row heights.", true, templatePath));
                            for (var column = 0; column < content.Header.Count; column++)
                            {
                                if (RequiresCode(content.Header[column]) && CountCodeSamples(rows[0].Elements<A.TableCell>().ElementAt(column)) != 1)
                                    AddCodeExemplarDiagnostic(diagnostics, "table header", column, CountCode(content.Header[column]), yolo, templatePath);
                                if (content.Rows.Any(row => RequiresCode(row[column])) && CountCodeSamples(rows[1].Elements<A.TableCell>().ElementAt(column)) != 1)
                                    AddCodeExemplarDiagnostic(diagnostics, "table body", column, content.Rows.Sum(row => CountCode(row[column])), yolo, templatePath);
                            }
                        }
                    }
                }
                if (region.Type == PresentationRegionType.Text)
                {
                    var shape = prototype.Slide.Descendants<P.Shape>().SingleOrDefault(candidate => candidate.TextBody?.InnerText.Contains(token, StringComparison.Ordinal) == true);
                    if (shape is null) diagnostics.Add(new("YADG-PPT-TEMPLATE-011", $"Text region '{region.Id}' must be an ordinary text shape.", true, templatePath));
                    else if (RequiresList(region.Blocks) && !shape.TextBody!.Elements<A.Paragraph>().Any(IsBulletParagraph)) diagnostics.Add(new("YADG-PPT-TEMPLATE-012", $"Text region '{region.Id}' has no bullet formatting example required by its content.", true, templatePath));
                    else if (RequiresHeading(region.Blocks) && !HasHeadingExample(shape.TextBody!, region.Id)) diagnostics.Add(new("YADG-PPT-TEMPLATE-014", $"Text region '{region.Id}' has no distinguishable content-heading formatting example.", true, templatePath));
                    if (shape is not null && RequiresCode(region.Blocks) && CountCodeSamples(shape) != 1)
                    {
                        var codeCount = CountCode(region.Blocks);
                        if (yolo) for (var i = 0; i < codeCount; i++) diagnostics.Add(new("YADG-YOLO-CODE-002", $"Text region '{region.Id}' has no unique {{codeInline}} formatting exemplar; fallback: preserve code as plain text.", false, templatePath) { IsDegradation = true });
                        else diagnostics.Add(new("YADG-PPT-CODE-001", $"Text region '{region.Id}' requires exactly one formatting exemplar run whose sole literal text is {{codeInline}}.", true, templatePath));
                    }
                }
            }
            if (texts.Any(text => AnyControl.IsMatch(text) && text is not "{{title}}" && !RegionMarker.IsMatch(text) && !SlideMarker.IsMatch(text)))
                diagnostics.Add(new("YADG-PPT-TEMPLATE-003", $"Layout '{slide.Layout}' contains a malformed or unknown control marker.", true, templatePath));
        }
    }

    private static SlidePart CloneSlide(PresentationPart presentationPart, SlidePart source)
    {
        var clone = presentationPart.AddNewPart<SlidePart>();
        using (var stream = source.GetStream(FileMode.Open, FileAccess.Read)) clone.FeedData(stream);
        foreach (var relationship in source.Parts) clone.AddPart(relationship.OpenXmlPart, relationship.RelationshipId);
        foreach (var relationship in source.ExternalRelationships) clone.AddExternalRelationship(relationship.RelationshipType, relationship.Uri, relationship.Id);
        foreach (var relationship in source.HyperlinkRelationships) clone.AddHyperlinkRelationship(relationship.Uri, relationship.IsExternal, relationship.Id);
        clone.Slide = (P.Slide)source.Slide.CloneNode(true);
        clone.Slide.Save();
        return clone;
    }

    private static void AuthorSlide(SlidePart slidePart, PresentationSlide slide, string root, List<Diagnostic> diagnostics)
    {
        foreach (var shape in slidePart.Slide.Descendants<P.Shape>().Where(shape => shape.TextBody is not null).ToArray())
        {
            var body = shape.TextBody!;
            var logical = body.InnerText;
            if (SlideMarker.IsMatch(logical)) { shape.Remove(); continue; }
            if (logical.Contains("{{title}}", StringComparison.Ordinal)) ReplaceShapeTitle(shape, slide, slidePart);
            foreach (var region in slide.Regions.Where(region => region.Type == PresentationRegionType.Text))
            {
                var token = $"{{{{{region.Id}}}}}";
                if (!logical.Contains(token, StringComparison.Ordinal)) continue;
                ReplaceTextRegion(shape, token, region.Blocks, diagnostics, slidePart);
            }

        }
        foreach (var region in slide.Regions.Where(region => region.Type == PresentationRegionType.Image))
        {
            var token = $"{{{{{region.Id}}}}}";
            var shape = slidePart.Slide.Descendants<P.Shape>().Single(candidate => candidate.TextBody?.InnerText.Contains(token, StringComparison.Ordinal) == true);
            var image = (PresentationImage)region.Blocks.Single();
            if (image.FailureReason is not null)
            {
                var identity = image.Id is null ? image.Path : $"Mermaid #{image.Id}";
                ReplaceShapeText(shape, $"YADG image unavailable: {identity} ({image.FailureReason})", null, token);
                continue;
            }
            string? error = null;
            var path = image.GeneratedSource is not null ? image.Path : ResolveImage(root, image, out error);
            if (path is null) { diagnostics.Add(new("YADG-PPT-ASSET-001", error!, true, image.SourcePath)); continue; }
            InsertPicture(slidePart, shape, path, image.Alt);
        }
        foreach (var region in slide.Regions.Where(region => region.Type == PresentationRegionType.Table))
        {
            var token = $"{{{{{region.Id}}}}}";
            var frame = slidePart.Slide.Descendants<P.GraphicFrame>().Single(candidate => candidate.Graphic?.GraphicData?.GetFirstChild<A.Table>()?.InnerText.Contains(token, StringComparison.Ordinal) == true);
            ReplaceTable(frame, (PresentationTable)region.Blocks.Single(), diagnostics, slidePart);
        }
    }

    private static void ReplaceShapeText(P.Shape shape, string text, A.Paragraph? exemplar, string marker = "{{title}}")
    {
        var body = shape.TextBody!;
        var markerParagraph = body.Elements<A.Paragraph>().FirstOrDefault(paragraph => paragraph.InnerText.Contains(marker, StringComparison.Ordinal));
        var source = exemplar ?? markerParagraph;
        if (source is null) return;
        var generated = NewParagraph(source, new[] { new PresentationText(text) });
        markerParagraph!.InsertAfterSelf(generated);
        markerParagraph.Remove();
    }

    private static void ReplaceShapeTitle(P.Shape shape, PresentationSlide slide, SlidePart slidePart)
    {
        var body = shape.TextBody!;
        var markerParagraph = body.Elements<A.Paragraph>().Single(paragraph => paragraph.InnerText.Contains("{{title}}", StringComparison.Ordinal));
        var codeProperties = CountCodeSamples(shape) == 1 ? shape.Descendants<A.Run>().FirstOrDefault(IsCodeSample)?.RunProperties : null;
        var generated = NewParagraph(markerParagraph, slide.TitleInlines, slidePart: slidePart, codeProperties: codeProperties);
        markerParagraph.InsertAfterSelf(generated);
        markerParagraph.Remove();
        RemoveCodeSamples(shape);
    }

    private static void ReplaceTextRegion(P.Shape shape, string token, IReadOnlyList<PresentationBlock> blocks, List<Diagnostic> diagnostics, SlidePart slidePart)
    {
        var body = shape.TextBody!;
        var paragraphs = body.Elements<A.Paragraph>().ToArray();
        var markerParagraph = paragraphs.SingleOrDefault(paragraph => paragraph.InnerText.Contains(token, StringComparison.Ordinal));
        if (markerParagraph is null) { diagnostics.Add(new("YADG-PPT-TEMPLATE-013", $"Unable to locate marker paragraph {token}.", true)); return; }
        var codeSample = CountCodeSamples(shape) == 1 ? shape.Descendants<A.Run>().FirstOrDefault(IsCodeSample)?.RunProperties : null;
        var samples = paragraphs.Where(paragraph => paragraph != markerParagraph && paragraph.InnerText.Length > 0 && !paragraph.Descendants<A.Run>().Any(IsCodeSample)).ToArray();
        var plain = samples.FirstOrDefault(paragraph => !IsBulletParagraph(paragraph));
        var unbulleted = samples.Where(paragraph => !IsBulletParagraph(paragraph)).OrderByDescending(ParagraphFontSize).ToArray();
        var heading = unbulleted.Length > 1 && ParagraphFontSize(unbulleted[0]) > ParagraphFontSize(unbulleted[^1]) ? unbulleted[0] : null;
        var bullet = samples.FirstOrDefault(IsBulletParagraph);
        var generated = new List<A.Paragraph>();
        foreach (var block in blocks)
        {
            switch (block)
            {
                case PresentationParagraph paragraph:
                    var sample = paragraph.HeadingLevel.HasValue ? heading : plain;
                    if (sample is null) { diagnostics.Add(new("YADG-PPT-TEMPLATE-014", paragraph.HeadingLevel.HasValue ? "Text region requires a content-heading formatting example." : "Text region requires a plain-paragraph formatting example.")); continue; }
                    generated.Add(NewParagraph(sample, paragraph.Inlines, slidePart: slidePart, codeProperties: codeSample));
                    break;
                case PresentationList list:
                    if (bullet is null) { diagnostics.Add(new("YADG-PPT-TEMPLATE-012", "Text region requires a bullet formatting example.")); continue; }
                    FlattenList(list, 0, generated, bullet, diagnostics, slidePart, codeSample);
                    break;
            }
        }
        OpenXmlElement? cursor = markerParagraph;
        foreach (var para in generated) { cursor.InsertAfterSelf(para); cursor = para; }
        foreach (var sample in samples) sample.Remove();
        markerParagraph.Remove();
        RemoveCodeSamples(shape);
    }

    private static void RemoveCodeSamples(P.Shape shape)
    {
        foreach (var run in shape.Descendants<A.Run>().Where(IsCodeSample).ToArray())
        {
            var paragraph = run.Ancestors<A.Paragraph>().FirstOrDefault();
            run.Remove();
            if (paragraph is not null && string.IsNullOrEmpty(paragraph.InnerText) && paragraph.Parent is not null) paragraph.Remove();
        }
    }

    private static void FlattenList(PresentationList list, int level, List<A.Paragraph> output, A.Paragraph exemplar, List<Diagnostic> diagnostics, SlidePart slidePart, A.RunProperties? codeProperties)
    {
        foreach (var item in list.Items)
        foreach (var block in item.Blocks)
        {
            if (block is PresentationParagraph paragraph) output.Add(NewParagraph(exemplar, paragraph.Inlines, level, slidePart, codeProperties));
            else if (block is PresentationList nested) FlattenList(nested, level + 1, output, exemplar, diagnostics, slidePart, codeProperties);
            else diagnostics.Add(new("YADG-PPT-MD-018", "Nested list content must be paragraph or list blocks."));
        }
    }

    private static A.Paragraph NewParagraph(A.Paragraph exemplar, IReadOnlyList<PresentationInline> inlines, int? listLevel = null, SlidePart? slidePart = null, A.RunProperties? codeProperties = null)
    {
        var result = (A.Paragraph)exemplar.CloneNode(true);
        foreach (var child in result.ChildElements.Where(child => child is A.Run or A.Break or A.Field).ToArray()) child.Remove();
        if (listLevel.HasValue)
        {
            var properties = result.GetFirstChild<A.ParagraphProperties>();
            if (properties is null) { properties = new A.ParagraphProperties(); result.PrependChild(properties); }
            properties.Level = (byte)Math.Clamp(listLevel.Value, 0, 8);
        }
        var sampleRunProperties = exemplar.Descendants<A.RunProperties>().FirstOrDefault();
        var end = result.GetFirstChild<A.EndParagraphRunProperties>();
        foreach (var inline in inlines) AppendInlineRuns(result, inline, sampleRunProperties, end, slidePart, codeProperties: codeProperties);
        return result;
    }

    private static void AppendInlineRuns(A.Paragraph paragraph, PresentationInline inline, A.RunProperties? baseProperties, A.EndParagraphRunProperties? end, SlidePart? slidePart, string? hyperlinkDestination = null, A.RunProperties? codeProperties = null)
    {
        switch (inline)
        {
            case PresentationText text: AddRun(text.Value, null); break;
            case PresentationCode code: AddRun(code.Value, codeProperties); break;
            case PresentationBreak: AddBreak(); break;
            case PresentationLink link:
                var destination = link.Destination;
                if (Uri.TryCreate(destination, UriKind.Absolute, out var target) && target.Scheme is "http" or "https" or "mailto" && slidePart is not null)
                    foreach (var child in link.Children) AppendInlineRuns(paragraph, child, baseProperties, end, slidePart, destination, codeProperties);
                else foreach (var child in link.Children) AppendInlineRuns(paragraph, child, baseProperties, end, slidePart, hyperlinkDestination, codeProperties);
                break;
            case PresentationEmphasis emphasis:
                var emphasized = baseProperties;
                if (emphasis.Strong) emphasized = WithBold(emphasized);
                if (emphasis.Italic || !emphasis.Strong) emphasized = WithItalic(emphasized);
                foreach (var child in emphasis.Children) AppendInlineRuns(paragraph, child, emphasized, end, slidePart, hyperlinkDestination, codeProperties);
                break;
            case PresentationStrike strike:
                foreach (var child in strike.Children) AppendInlineRuns(paragraph, child, WithStrike(baseProperties), end, slidePart, hyperlinkDestination, codeProperties);
                break;
        }
        void AddRun(string value, A.RunProperties? extra)
        {
            var properties = (A.RunProperties?)(extra ?? baseProperties?.CloneNode(true)) ?? new A.RunProperties();
            if (extra is not null && baseProperties is not null)
            {
                if (baseProperties.Bold is not null) properties.Bold = baseProperties.Bold;
                if (baseProperties.Italic is not null) properties.Italic = baseProperties.Italic;
                if (baseProperties.Strike is not null) properties.Strike = baseProperties.Strike;
            }
            if (hyperlinkDestination is not null && slidePart is not null && Uri.TryCreate(hyperlinkDestination, UriKind.Absolute, out var target))
            {
                var relationship = slidePart.AddHyperlinkRelationship(target, true);
                properties.AppendChild(new A.HyperlinkOnClick { Id = relationship.Id });
            }
            var run = new A.Run(properties, new A.Text(value));
            if (end is null) paragraph.AppendChild(run); else paragraph.InsertBefore(run, end);
        }
        void AddBreak()
        {
            var br = new A.Break();
            if (end is null) paragraph.AppendChild(br); else paragraph.InsertBefore(br, end);
        }
    }

    private static A.RunProperties WithBold(A.RunProperties? source) { var result = source is null ? new A.RunProperties() : (A.RunProperties)source.CloneNode(true); result.Bold = true; return result; }
    private static A.RunProperties WithItalic(A.RunProperties? source) { var result = source is null ? new A.RunProperties() : (A.RunProperties)source.CloneNode(true); result.Italic = true; return result; }
    private static A.RunProperties WithStrike(A.RunProperties? source) { var result = source is null ? new A.RunProperties() : (A.RunProperties)source.CloneNode(true); result.Strike = A.TextStrikeValues.SingleStrike; return result; }
    private static bool IsBulletParagraph(A.Paragraph paragraph)
    {
        if (paragraph.ParagraphProperties?.Elements<A.CharacterBullet>().Any() == true || paragraph.ParagraphProperties?.Elements<A.AutoNumberedBullet>().Any() == true) return true;
        var nativePlaceholder = paragraph.Ancestors<P.Shape>().Any(shape => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is not null);
        return nativePlaceholder && paragraph.ParagraphProperties?.Level is not null;
    }
    private static bool IsCodeSample(A.Run run) => run.Elements<A.Text>().Count() == 1 && run.Elements<A.Text>().Single().Text == "{{codeInline}}";
    private static int CountCodeSamples(OpenXmlElement element) => element.Descendants<A.Run>().Count(IsCodeSample);
    private static void AddCodeExemplarDiagnostic(List<Diagnostic> diagnostics, string scope, int column, int count, bool yolo, string templatePath)
    {
        if (yolo) for (var i = 0; i < count; i++) diagnostics.Add(new("YADG-YOLO-CODE-002", $"{scope} cell {column + 1} has no unique {{codeInline}} formatting exemplar; fallback: preserve code as plain text.", false, templatePath) { IsDegradation = true });
        else diagnostics.Add(new("YADG-PPT-CODE-002", $"{scope} cell {column + 1} requires exactly one {{codeInline}} formatting exemplar.", true, templatePath));
    }
    private static bool RequiresCode(IEnumerable<PresentationBlock> blocks) => blocks.Any(block => block switch
    {
        PresentationParagraph paragraph => RequiresCode(paragraph.Inlines),
        PresentationList list => list.Items.SelectMany(item => item.Blocks).Any(nested => nested is PresentationParagraph child && RequiresCode(child.Inlines) || nested is PresentationList inner && RequiresCode(new PresentationBlock[] { inner })),
        PresentationTable table => table.Header.Any(RequiresCode) || table.Rows.SelectMany(row => row).Any(RequiresCode),
        _ => false
    });
    private static bool RequiresCode(IEnumerable<PresentationInline> inlines) => inlines.Any(inline => inline switch
    {
        PresentationCode => true,
        PresentationEmphasis emphasis => RequiresCode(emphasis.Children),
        PresentationStrike strike => RequiresCode(strike.Children),
        PresentationLink link => RequiresCode(link.Children),
        _ => false
    });
    private static int CountCode(IEnumerable<PresentationBlock> blocks) => blocks.Sum(block => block switch
    {
        PresentationParagraph paragraph => CountCode(paragraph.Inlines),
        PresentationList list => list.Items.SelectMany(item => item.Blocks).Sum(nested => nested is PresentationParagraph child ? CountCode(child.Inlines) : nested is PresentationList inner ? CountCode(new PresentationBlock[] { inner }) : 0),
        _ => 0
    });
    private static int CountCode(IEnumerable<PresentationInline> inlines) => inlines.Sum(inline => inline switch
    {
        PresentationCode => 1,
        PresentationEmphasis emphasis => CountCode(emphasis.Children),
        PresentationStrike strike => CountCode(strike.Children),
        PresentationLink link => CountCode(link.Children),
        _ => 0
    });
    private static bool RequiresList(IEnumerable<PresentationBlock> blocks) => blocks.OfType<PresentationList>().Any();
    private static bool RequiresHeading(IEnumerable<PresentationBlock> blocks) => blocks.OfType<PresentationParagraph>().Any(block => block.HeadingLevel.HasValue);
    private static bool HasHeadingExample(P.TextBody body, string marker)
    {
        var samples = body.Elements<A.Paragraph>().Where(paragraph => !paragraph.InnerText.Contains($"{{{{{marker}}}}}", StringComparison.Ordinal) && paragraph.InnerText.Length > 0 && !IsBulletParagraph(paragraph)).OrderByDescending(ParagraphFontSize).ToArray();
        return samples.Length > 1 && ParagraphFontSize(samples[0]) > ParagraphFontSize(samples[^1]);
    }
    private static int ParagraphFontSize(A.Paragraph paragraph) => paragraph.Descendants<A.RunProperties>().Select(properties => (int?)properties.FontSize?.Value).Max() ?? 0;

    private static void InsertPicture(SlidePart part, P.Shape sourceShape, string path, string alt)
    {
        var xfrm = sourceShape.ShapeProperties?.Transform2D ?? throw new InvalidDataException("Image prototype shape has no transform bounds.");
        var offset = xfrm.Offset!; var extents = xfrm.Extents!;
        var (pixelWidth, pixelHeight, isJpeg) = ImageDimensions(path);
        var scale = Math.Min((double)extents.Cx!.Value / pixelWidth, (double)extents.Cy!.Value / pixelHeight);
        var width = (long)Math.Round(pixelWidth * scale); var height = (long)Math.Round(pixelHeight * scale);
        var x = offset.X!.Value + (extents.Cx.Value - width) / 2; var y = offset.Y!.Value + (extents.Cy.Value - height) / 2;
        var imagePart = part.AddImagePart(isJpeg ? ImagePartType.Jpeg : ImagePartType.Png);
        using (var stream = File.OpenRead(path)) imagePart.FeedData(stream);
        var id = part.GetIdOfPart(imagePart);
        var pictureId = (part.Slide.Descendants<P.NonVisualDrawingProperties>().Select(properties => properties.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max()) + 1;
        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = pictureId, Name = Path.GetFileName(path), Description = alt },
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(new A.Blip { Embed = id }, new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = width, Cy = height }), new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        sourceShape.InsertAfterSelf(picture);
        sourceShape.Remove();
    }

    private static void ReplaceTable(P.GraphicFrame frame, PresentationTable content, List<Diagnostic> diagnostics, SlidePart slidePart)
    {
        var table = frame.Graphic?.GraphicData?.GetFirstChild<A.Table>() ?? throw new InvalidDataException("Table region has no native table.");
        var rows = table.Elements<A.TableRow>().ToArray();
        if (rows.Length != 2) { diagnostics.Add(new("YADG-PPT-TEMPLATE-015", "Table prototype must contain exactly one header example row and one body example row.")); return; }
        var header = rows[0]; var body = rows[1];
        var headerCells = header.Elements<A.TableCell>().ToArray(); var bodyCells = body.Elements<A.TableCell>().ToArray();
        if (headerCells.Length != content.Header.Count || bodyCells.Length != content.Header.Count || content.Rows.Any(row => row.Count != content.Header.Count))
        { diagnostics.Add(new("YADG-PPT-TABLE-001", "Source table column count must match the native prototype table.")); return; }
        if (table.Descendants<A.TableCell>().Any(cell => (cell.RowSpan?.Value ?? 1) > 1 || (cell.GridSpan?.Value ?? 1) > 1))
        { diagnostics.Add(new("YADG-PPT-TEMPLATE-016", "Merged prototype table cells are unsupported.")); return; }
        var shapeExtent = frame.Transform?.Extents?.Cy?.Value;
        var headerHeight = header.Height?.Value ?? 0;
        var bodyHeight = body.Height?.Value ?? 0;
        if (shapeExtent is > 0 && headerHeight > 0 && bodyHeight > 0 && headerHeight + bodyHeight * content.Rows.Count > shapeExtent.Value)
        { diagnostics.Add(new("YADG-PPT-TABLE-002", "Table rows cannot fit within the prototype's fixed structural height while preserving example row heights.")); return; }
        for (var column = 0; column < headerCells.Length; column++) ReplaceCell(headerCells[column], content.Header[column], slidePart);
        for (var column = 0; column < bodyCells.Length; column++) ReplaceCell(bodyCells[column], content.Rows[0][column], slidePart);
        for (var rowIndex = 1; rowIndex < content.Rows.Count; rowIndex++)
        {
            var clone = (A.TableRow)body.CloneNode(true);
            var cells = clone.Elements<A.TableCell>().ToArray();
            for (var column = 0; column < cells.Length; column++) ReplaceCell(cells[column], content.Rows[rowIndex][column], slidePart);
            body.InsertAfterSelf(clone);
            body = clone;
        }
    }

    private static void ReplaceCell(A.TableCell cell, IReadOnlyList<PresentationInline> content, SlidePart slidePart)
    {
        var old = cell.Elements<A.TextBody>().SingleOrDefault();
        if (old is null) { old = new A.TextBody(new A.BodyProperties(), new A.ListStyle()); cell.AppendChild(old); }
        var exemplar = old.Elements<A.Paragraph>().FirstOrDefault() ?? new A.Paragraph(new A.ParagraphProperties());
        var codeProperties = CountCodeSamples(cell) == 1 ? cell.Descendants<A.Run>().FirstOrDefault(IsCodeSample)?.RunProperties : null;
        var generated = NewParagraph(exemplar, content, slidePart: slidePart, codeProperties: codeProperties);
        old.RemoveAllChildren<A.Paragraph>();
        old.AppendChild(generated);
    }

    private static string? ResolveImage(string root, PresentationImage image, out string? error)
    {
        error = null;
        try
        {
            if (Uri.TryCreate(image.Path, UriKind.Absolute, out _) || image.Path.Contains('*') || image.Path.Contains('?')) { error = $"Image must use a local workspace-relative path: '{image.Path}'."; return null; }
            var full = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(image.SourcePath) ?? root, image.Path));
            var relative = Path.GetRelativePath(Path.GetFullPath(root), full);
            if (relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal) || Path.IsPathRooted(relative)) { error = $"Image must remain lexically inside the effective workspace: '{image.Path}'."; return null; }
            if (!File.Exists(full)) { error = $"Image asset does not exist: '{image.Path}'."; return null; }
            if (Path.GetExtension(full).ToLowerInvariant() is not ".png" and not ".jpg" and not ".jpeg") { error = "Only PNG and JPEG images are supported."; return null; }
            _ = ImageDimensions(full);
            return full;
        }
        catch (Exception ex) { error = $"Unable to validate image: {ex.Message}"; return null; }
    }

    private static (int Width, int Height, bool IsJpeg) ImageDimensions(string path)
    {
        var bytes = File.ReadAllBytes(path);
        if (Path.GetExtension(path).Equals(".png", StringComparison.OrdinalIgnoreCase) && bytes.Length >= 24 && bytes[0] == 137 && bytes[1] == 80 && bytes[2] == 78 && bytes[3] == 71)
            return (ReadInt(bytes, 16), ReadInt(bytes, 20), false);
        if (bytes.Length > 4 && bytes[0] == 0xff && bytes[1] == 0xd8)
        {
            var i = 2;
            while (i + 9 < bytes.Length)
            {
                if (bytes[i++] != 0xff) continue;
                var marker = bytes[i++];
                if (marker is 0xd8 or 0xd9) continue;
                var length = bytes[i] << 8 | bytes[i + 1];
                if (length < 2 || i + length > bytes.Length) break;
                if (marker is >= 0xc0 and <= 0xc3 or >= 0xc5 and <= 0xc7 or >= 0xc9 and <= 0xcb or >= 0xcd and <= 0xcf)
                    return (bytes[i + 3] << 8 | bytes[i + 4], bytes[i + 5] << 8 | bytes[i + 6], true);
                i += length;
            }
        }
        throw new InvalidDataException($"Image is not a decodable PNG or JPEG: '{path}'.");
    }

    private static int ReadInt(byte[] bytes, int offset) => bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3];
}
