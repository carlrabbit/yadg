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

public sealed record PresentationAnalysis(PresentationDocument Document, IReadOnlyList<Diagnostic> Diagnostics)
{
    public bool IsValid => Diagnostics.All(diagnostic => !diagnostic.IsError) && Document.IsValid;
}

public static class PresentationAuthoring
{
    private static readonly Regex SlideMarker = new(@"^\{\{slide:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$", RegexOptions.Compiled);
    private static readonly Regex RegionMarker = new(@"^\{\{(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$", RegexOptions.Compiled);
    private static readonly Regex AnyControl = new(@"\{\{[^{}]*\}\}", RegexOptions.Compiled);

    public static PresentationAnalysis Analyze(string templatePath, IEnumerable<PresentationSource> sources, string effectiveRoot, string inlineCodePolicy = "style")
    {
        var document = PresentationMarkdownParser.Parse(sources, inlineCodePolicy);
        var diagnostics = document.Diagnostics.ToList();
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
            ValidateSelectedPrototypes(presentation.PresentationPart!, document, prototypes, diagnostics, templatePath);
        }
        catch (Exception ex) { diagnostics.Add(new("YADG-PPT-TEMPLATE-002", $"Unable to inspect PowerPoint template: {ex.Message}", true, templatePath)); return new(document, diagnostics); }

        foreach (var slide in document.Slides)
        foreach (var region in slide.Regions.Where(region => region.Type == PresentationRegionType.Image))
        foreach (var image in region.Blocks.OfType<PresentationImage>())
        {
            var resolved = ResolveImage(effectiveRoot, image, out var error);
            if (resolved is null) diagnostics.Add(new("YADG-PPT-ASSET-001", error!, true, image.SourcePath));
        }
        return new(document, diagnostics);
    }

    public static IReadOnlyList<Diagnostic> Build(string templatePath, IEnumerable<PresentationSource> sources, string effectiveRoot, string outputPath, string inlineCodePolicy = "style")
    {
        var sourceArray = sources.ToArray();
        var analysis = Analyze(templatePath, sourceArray, effectiveRoot, inlineCodePolicy);
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
        return Array.Empty<Diagnostic>();
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

    private static void ValidateSelectedPrototypes(PresentationPart package, PresentationDocument document, IReadOnlyDictionary<string, SlidePart> prototypes, List<Diagnostic> diagnostics, string templatePath)
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
                        }
                    }
                }
                if (region.Type == PresentationRegionType.Text)
                {
                    var shape = prototype.Slide.Descendants<P.Shape>().SingleOrDefault(candidate => candidate.TextBody?.InnerText.Contains(token, StringComparison.Ordinal) == true);
                    if (shape is null) diagnostics.Add(new("YADG-PPT-TEMPLATE-011", $"Text region '{region.Id}' must be an ordinary text shape.", true, templatePath));
                    else if (RequiresList(region.Blocks) && !shape.TextBody!.Elements<A.Paragraph>().Any(IsBulletParagraph)) diagnostics.Add(new("YADG-PPT-TEMPLATE-012", $"Text region '{region.Id}' has no bullet formatting example required by its content.", true, templatePath));
                    else if (RequiresHeading(region.Blocks) && !HasHeadingExample(shape.TextBody!, region.Id)) diagnostics.Add(new("YADG-PPT-TEMPLATE-014", $"Text region '{region.Id}' has no distinguishable content-heading formatting example.", true, templatePath));
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
            if (logical == "{{title}}") ReplaceShapeText(shape, slide.Title, null);
            foreach (var region in slide.Regions.Where(region => region.Type == PresentationRegionType.Text))
            {
                var token = $"{{{{{region.Id}}}}}";
                if (!logical.Contains(token, StringComparison.Ordinal)) continue;
                ReplaceTextRegion(shape, token, region.Blocks, diagnostics);
            }
        }
        foreach (var region in slide.Regions.Where(region => region.Type == PresentationRegionType.Image))
        {
            var token = $"{{{{{region.Id}}}}}";
            var shape = slidePart.Slide.Descendants<P.Shape>().Single(candidate => candidate.TextBody?.InnerText.Contains(token, StringComparison.Ordinal) == true);
            var image = (PresentationImage)region.Blocks.Single();
            var path = ResolveImage(root, image, out var error);
            if (path is null) { diagnostics.Add(new("YADG-PPT-ASSET-001", error!, true, image.SourcePath)); continue; }
            InsertPicture(slidePart, shape, path, image.Alt);
        }
        foreach (var region in slide.Regions.Where(region => region.Type == PresentationRegionType.Table))
        {
            var token = $"{{{{{region.Id}}}}}";
            var frame = slidePart.Slide.Descendants<P.GraphicFrame>().Single(candidate => candidate.Graphic?.GraphicData?.GetFirstChild<A.Table>()?.InnerText.Contains(token, StringComparison.Ordinal) == true);
            ReplaceTable(frame, (PresentationTable)region.Blocks.Single(), diagnostics);
        }
    }

    private static void ReplaceShapeText(P.Shape shape, string text, A.Paragraph? exemplar)
    {
        var body = shape.TextBody!;
        var markerParagraph = body.Elements<A.Paragraph>().FirstOrDefault(paragraph => paragraph.InnerText.Contains("{{title}}", StringComparison.Ordinal));
        var source = exemplar ?? markerParagraph;
        if (source is null) return;
        var generated = NewParagraph(source, new[] { new PresentationText(text) });
        markerParagraph!.InsertAfterSelf(generated);
        markerParagraph.Remove();
    }

    private static void ReplaceTextRegion(P.Shape shape, string token, IReadOnlyList<PresentationBlock> blocks, List<Diagnostic> diagnostics)
    {
        var body = shape.TextBody!;
        var paragraphs = body.Elements<A.Paragraph>().ToArray();
        var markerParagraph = paragraphs.SingleOrDefault(paragraph => paragraph.InnerText.Contains(token, StringComparison.Ordinal));
        if (markerParagraph is null) { diagnostics.Add(new("YADG-PPT-TEMPLATE-013", $"Unable to locate marker paragraph {token}.", true)); return; }
        var samples = paragraphs.Where(paragraph => paragraph != markerParagraph && paragraph.InnerText.Length > 0).ToArray();
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
                    generated.Add(NewParagraph(sample, paragraph.Inlines));
                    break;
                case PresentationList list:
                    if (bullet is null) { diagnostics.Add(new("YADG-PPT-TEMPLATE-012", "Text region requires a bullet formatting example.")); continue; }
                    FlattenList(list, 0, generated, bullet, diagnostics);
                    break;
            }
        }
        OpenXmlElement? cursor = markerParagraph;
        foreach (var para in generated) { cursor.InsertAfterSelf(para); cursor = para; }
        foreach (var sample in samples) sample.Remove();
        markerParagraph.Remove();
    }

    private static void FlattenList(PresentationList list, int level, List<A.Paragraph> output, A.Paragraph exemplar, List<Diagnostic> diagnostics)
    {
        foreach (var item in list.Items)
        foreach (var block in item.Blocks)
        {
            if (block is PresentationParagraph paragraph) output.Add(NewParagraph(exemplar, paragraph.Inlines, level));
            else if (block is PresentationList nested) FlattenList(nested, level + 1, output, exemplar, diagnostics);
            else diagnostics.Add(new("YADG-PPT-MD-018", "Nested list content must be paragraph or list blocks."));
        }
    }

    private static A.Paragraph NewParagraph(A.Paragraph exemplar, IReadOnlyList<PresentationInline> inlines, int? listLevel = null)
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
        foreach (var inline in inlines) AppendInlineRuns(result, inline, sampleRunProperties, end);
        return result;
    }

    private static void AppendInlineRuns(A.Paragraph paragraph, PresentationInline inline, A.RunProperties? baseProperties, A.EndParagraphRunProperties? end)
    {
        switch (inline)
        {
            case PresentationText text: AddRun(text.Value, null); break;
            case PresentationCode code: AddRun(code.Value, null); break;
            case PresentationBreak: AddBreak(); break;
            case PresentationLink link:
                foreach (var child in link.Children) AppendInlineRuns(paragraph, child, baseProperties, end);
                break;
            case PresentationEmphasis emphasis:
                foreach (var child in emphasis.Children) AppendInlineRuns(paragraph, child, emphasis.Strong ? WithBold(baseProperties) : WithItalic(baseProperties), end);
                break;
        }
        void AddRun(string value, A.RunProperties? extra)
        {
            var properties = (A.RunProperties?)(extra ?? baseProperties?.CloneNode(true)) ?? new A.RunProperties();
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
    private static bool IsBulletParagraph(A.Paragraph paragraph) => paragraph.ParagraphProperties?.Elements<A.CharacterBullet>().Any() == true || paragraph.ParagraphProperties?.Elements<A.AutoNumberedBullet>().Any() == true;
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

    private static void ReplaceTable(P.GraphicFrame frame, PresentationTable content, List<Diagnostic> diagnostics)
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
        for (var column = 0; column < headerCells.Length; column++) ReplaceCell(headerCells[column], content.Header[column]);
        for (var column = 0; column < bodyCells.Length; column++) ReplaceCell(bodyCells[column], content.Rows[0][column]);
        for (var rowIndex = 1; rowIndex < content.Rows.Count; rowIndex++)
        {
            var clone = (A.TableRow)body.CloneNode(true);
            var cells = clone.Elements<A.TableCell>().ToArray();
            for (var column = 0; column < cells.Length; column++) ReplaceCell(cells[column], content.Rows[rowIndex][column]);
            body.InsertAfterSelf(clone);
            body = clone;
        }
    }

    private static void ReplaceCell(A.TableCell cell, IReadOnlyList<PresentationInline> content)
    {
        var old = cell.Elements<A.TextBody>().SingleOrDefault();
        if (old is null) { old = new A.TextBody(new A.BodyProperties(), new A.ListStyle()); cell.AppendChild(old); }
        var exemplar = old.Elements<A.Paragraph>().FirstOrDefault() ?? new A.Paragraph(new A.ParagraphProperties());
        var generated = NewParagraph(exemplar, content);
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
