using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Yadg.Core;

namespace Yadg.Word;

public enum PlacementKind { Section, Table, Figure }
public sealed record TemplatePlacement(Paragraph Anchor, PlacementKind Kind, string Id, SectionReference? SectionReference = null);
public sealed record TemplateAnalysis(IReadOnlyList<TemplatePlacement> Placements, IReadOnlyList<Diagnostic> Diagnostics)
{
    public bool IsValid => Diagnostics.All(d => !d.IsError);
}

public sealed class WordAuthoringException(Diagnostic diagnostic) : Exception(diagnostic.ToString())
{
    public Diagnostic Diagnostic { get; } = diagnostic;
}

public static class WordAuthoring
{
    private static readonly Regex TagLike = new(@"\{\{[^{}]*\}\}", RegexOptions.Compiled);
    private static readonly Regex DirectTag = new(@"^\{\{(table|figure):([A-Za-z][A-Za-z0-9_-]*)\}\}$", RegexOptions.Compiled);

    public static TemplateAnalysis Analyze(string templatePath, YadgDocument model)
    {
        using var document = WordprocessingDocument.Open(templatePath, false);
        return AnalyzeOpen(document, templatePath, model);
    }

    private static TemplateAnalysis AnalyzeOpen(WordprocessingDocument document, string templatePath, YadgDocument model)
    {
        var diagnostics = new List<Diagnostic>();
        var placements = new List<TemplatePlacement>();
        var main = document.MainDocumentPart;
        if (main?.Document.Body is null) return new(placements, new[] { new Diagnostic("YADG-WORD-001", "DOCX has no main document body.", true, templatePath) });
        var styles = main.StyleDefinitionsPart?.Styles;
        var numbering = main.NumberingDefinitionsPart?.Numbering;
        AnalyzeParagraphs(main.Document.Body.Descendants<Paragraph>(), templatePath, model, styles, numbering, placements, diagnostics, true);
        foreach (var header in main.HeaderParts) AnalyzeParagraphs(header.Header.Descendants<Paragraph>(), templatePath, model, styles, numbering, placements, diagnostics, false);
        foreach (var footer in main.FooterParts) AnalyzeParagraphs(footer.Footer.Descendants<Paragraph>(), templatePath, model, styles, numbering, placements, diagnostics, false);
        foreach (var duplicate in placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure).GroupBy(p => (p.Kind, p.Id)).Where(g => g.Count() > 1))
            diagnostics.Add(new("YADG-PLACEMENT-002", $"Structured object '{duplicate.Key.Id}' has more than one direct placement in this template.", true, templatePath));
        return new(placements, diagnostics);
    }

    public static void Author(string templatePath, string outputPath, YadgDocument model)
    {
        using (var analysisDocument = WordprocessingDocument.Open(templatePath, false))
        {
            var analysis = AnalyzeOpen(analysisDocument, templatePath, model);
            if (!analysis.IsValid) throw new WordAuthoringException(analysis.Diagnostics.First(d => d.IsError));
        }
        File.Copy(templatePath, outputPath, true);
        using var document = WordprocessingDocument.Open(outputPath, true);
        var analysisOutput = AnalyzeOpen(document, outputPath, model);
        if (!analysisOutput.IsValid) throw new WordAuthoringException(analysisOutput.Diagnostics.First(d => d.IsError));
        var direct = analysisOutput.Placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure).Select(p => (p.Kind, p.Id)).ToHashSet();
        foreach (var placement in analysisOutput.Placements.Reverse())
        {
            var blocks = placement.Kind == PlacementKind.Section
                ? MarkdownDocumentParser.Resolve(model, placement.SectionReference!, outputPath).Blocks!
                : new[] { placement.Kind == PlacementKind.Table ? (YadgBlock)model.FindTable(placement.Id)! : model.FindFigure(placement.Id)! };
            var elements = new List<OpenXmlElement>();
            foreach (var block in blocks)
            {
                if (placement.Kind == PlacementKind.Section && block is YadgTable table && direct.Contains((PlacementKind.Table, table.Id))) continue;
                if (placement.Kind == PlacementKind.Section && block is YadgFigure figure && direct.Contains((PlacementKind.Figure, figure.Id))) continue;
                elements.AddRange(CreateElements(block, placement.Anchor, document));
            }
            foreach (var element in elements) placement.Anchor.InsertBeforeSelf(element);
            placement.Anchor.Remove();
        }
        document.MainDocumentPart!.Document.Save();
    }

    private static void AnalyzeParagraphs(IEnumerable<Paragraph> paragraphs, string templatePath, YadgDocument model, Styles? styles, Numbering? numbering, List<TemplatePlacement> placements, List<Diagnostic> diagnostics, bool mainBody)
    {
        foreach (var paragraph in paragraphs)
        {
            var logical = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
            var matches = TagLike.Matches(logical).Cast<Match>().ToArray();
            if (matches.Length == 0) continue;
            if (!mainBody || paragraph.Parent is not Body || paragraph.Ancestors<Table>().Any())
            { diagnostics.Add(new("YADG-WORD-LOCATION", "YADG tags are supported only as standalone paragraphs in the main document body.", true, templatePath)); continue; }
            foreach (var match in matches)
            {
                if (!string.Equals(logical.Trim(), match.Value, StringComparison.Ordinal)) { diagnostics.Add(new("YADG-WORD-BLOCK", $"Block tag '{match.Value}' must be the only non-whitespace paragraph content.", true, templatePath)); continue; }
                var direct = DirectTag.Match(match.Value);
                if (direct.Success)
                {
                    var kind = direct.Groups[1].Value == "table" ? PlacementKind.Table : PlacementKind.Figure; var id = direct.Groups[2].Value;
                    if (kind == PlacementKind.Table && model.FindTable(id) is null) { diagnostics.Add(new("YADG-REF-003", $"Tag '{match.Value}' does not resolve to a table.", true, templatePath)); continue; }
                    if (kind == PlacementKind.Figure && model.FindFigure(id) is null) { diagnostics.Add(new("YADG-REF-003", $"Tag '{match.Value}' does not resolve to a figure.", true, templatePath)); continue; }
                    if (kind == PlacementKind.Table) ValidateTable(model.FindTable(id)!, styles, templatePath, diagnostics);
                    else ValidateFigure(model.FindFigure(id)!, styles, paragraph, templatePath, diagnostics);
                    placements.Add(new(paragraph, kind, id)); continue;
                }
                if (SectionReference.TryParseTag(match.Value, out var sectionReference, out _))
                {
                    var resolved = MarkdownDocumentParser.Resolve(model, sectionReference!, templatePath);
                    if (resolved.Diagnostic is not null) { diagnostics.Add(resolved.Diagnostic); continue; }
                    ValidateBlocks(resolved.Blocks!, styles, numbering, paragraph, templatePath, diagnostics);
                    placements.Add(new(paragraph, PlacementKind.Section, sectionReference!.Id, sectionReference));
                }
                else if (match.Value.StartsWith("{{value:", StringComparison.Ordinal)) diagnostics.Add(new("YADG-VALUE-UNSUPPORTED", $"Reserved value tag '{match.Value}' is not supported in M0003.", true, templatePath));
                else diagnostics.Add(new("YADG-TAG-001", $"Malformed or unsupported tag '{match.Value}'.", true, templatePath));
            }
        }
    }

    private static void ValidateBlocks(IEnumerable<YadgBlock> blocks, Styles? styles, Numbering? numbering, Paragraph anchor, string location, List<Diagnostic> diagnostics)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case YadgHeading heading when !HasStyle(styles, $"Heading{heading.Level}"): diagnostics.Add(new("YADG-WORD-STYLE", $"Missing required Word style 'Heading{heading.Level}'.", true, location)); break;
                case YadgList list: ValidateList(list, styles, numbering, location, diagnostics); break;
                case YadgTable table: ValidateTable(table, styles, location, diagnostics); break;
                case YadgFigure figure: ValidateFigure(figure, styles, anchor, location, diagnostics); break;
            }
        }
    }

    private static void ValidateList(YadgList list, Styles? styles, Numbering? numbering, string location, List<Diagnostic> diagnostics)
    {
        var style = list.Ordered ? "ListNumber" : "ListBullet";
        if (!HasStyle(styles, style)) { diagnostics.Add(new("YADG-WORD-LIST", $"Missing required list style '{style}'.", true, location)); return; }
        var numId = ResolveNumberingId(styles!, style);
        if (numId is null || numbering is null || !numbering.Descendants<NumberingInstance>().Any(n => n.NumberID?.Value == numId)) diagnostics.Add(new("YADG-WORD-LIST", $"List style '{style}' does not resolve to a valid numbering definition.", true, location));
    }

    private static int? ResolveNumberingId(Styles styles, string styleId)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (seen.Add(styleId))
        {
            var style = styles.Descendants<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style is null) return null;
            var direct = style.Descendants<NumberingProperties>().FirstOrDefault()?.NumberingId?.Val?.Value;
            if (direct is not null) return direct;
            styleId = style.BasedOn?.Val?.Value ?? string.Empty;
            if (styleId.Length == 0) return null;
        }
        return null;
    }

    private static void ValidateTable(YadgTable table, Styles? styles, string location, List<Diagnostic> diagnostics)
    { if (!HasStyle(styles, "TableGrid")) diagnostics.Add(new("YADG-WORD-TABLE", "Missing required Word table style 'TableGrid'.", true, location)); }

    private static void ValidateFigure(YadgFigure figure, Styles? styles, Paragraph anchor, string location, List<Diagnostic> diagnostics)
    {
        if (figure.Asset is null) { diagnostics.Add(new("YADG-FIGURE-008", $"Figure '{figure.Id}' has no validated image asset.", true, location)); return; }
        if (!string.IsNullOrEmpty(figure.AltText) && !HasStyle(styles, "Caption")) diagnostics.Add(new("YADG-WORD-STYLE", "Missing required Word style 'Caption'.", true, location));
        if (EffectiveWidth(anchor) <= 0) diagnostics.Add(new("YADG-FIGURE-009", $"Cannot determine effective text width for figure '{figure.Id}'.", true, location));
    }

    private static bool HasStyle(Styles? styles, string id) => styles?.Descendants<Style>().Any(s => s.StyleId?.Value == id) == true;

    private static IEnumerable<OpenXmlElement> CreateElements(YadgBlock block, Paragraph anchor, WordprocessingDocument document)
    {
        switch (block)
        {
            case YadgHeading heading: yield return CreateParagraph(heading, anchor); break;
            case YadgParagraph paragraph: yield return CreateParagraph(paragraph, anchor); break;
            case YadgList list:
                foreach (var item in list.Items) yield return CreateListParagraph(item, list.Ordered); break;
            case YadgTable table: yield return CreateTable(table); break;
            case YadgFigure figure:
                var imageParagraph = CreateFigureParagraph(figure, anchor, document); yield return imageParagraph;
                if (!string.IsNullOrEmpty(figure.AltText)) yield return new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }), new Run(new Text(figure.AltText) { Space = SpaceProcessingModeValues.Preserve }));
                break;
        }
    }

    private static Paragraph CreateParagraph(YadgBlock block, Paragraph anchor)
    {
        var paragraph = new Paragraph();
        if (block is YadgHeading heading) paragraph.Append(new ParagraphProperties(new ParagraphStyleId { Val = $"Heading{heading.Level}" }));
        else if (anchor.ParagraphProperties is not null) paragraph.Append(anchor.ParagraphProperties.CloneNode(true));
        AppendInlines(paragraph, block is YadgHeading h ? new[] { new YadgText(h.Text) } : ((YadgParagraph)block).Inlines);
        return paragraph;
    }

    private static Paragraph CreateListParagraph(YadgListItem item, bool ordered)
    {
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = ordered ? "ListNumber" : "ListBullet" })); AppendInlines(paragraph, item.Inlines); return paragraph;
    }

    private static Table CreateTable(YadgTable table)
    {
        var result = new Table(new TableProperties(new TableStyle { Val = "TableGrid" }, new TableLayout { Type = TableLayoutValues.Autofit }, new TableLook { FirstRow = OnOffValue.FromBoolean(true) }));
        result.Append(new TableGrid(Enumerable.Range(0, table.Header.Count).Select(_ => new GridColumn())));
        var header = new TableRow(new TableRowProperties(new TableHeader())); foreach (var cell in table.Header) header.Append(CreateCell(cell)); result.Append(header);
        foreach (var row in table.Rows) { var tableRow = new TableRow(); foreach (var cell in row) tableRow.Append(CreateCell(cell)); result.Append(tableRow); }
        return result;
    }

    private static TableCell CreateCell(IReadOnlyList<YadgInline> inlines) { var paragraph = new Paragraph(); AppendInlines(paragraph, inlines); return new TableCell(paragraph); }

    private static Paragraph CreateFigureParagraph(YadgFigure figure, Paragraph anchor, WordprocessingDocument document)
    {
        var asset = figure.Asset!; var main = document.MainDocumentPart!; var partType = asset.Format == "PNG" ? ImagePartType.Png : ImagePartType.Jpeg; var part = main.AddImagePart(partType);
        using (var stream = File.OpenRead(asset.FullPath)) part.FeedData(stream);
        var relation = main.GetIdOfPart(part); var width = Math.Min((long)asset.WidthPixels * 9525, EffectiveWidth(anchor)); var height = Math.Max(1, (long)asset.HeightPixels * 9525 * width / ((long)asset.WidthPixels * 9525));
        var drawing = new Drawing(new DW.Inline(new DW.Extent { Cx = width, Cy = height }, new DW.DocProperties { Id = 1U, Name = figure.Id }, new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }), new A.Graphic(new A.GraphicData(new PIC.Picture(new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties { Id = 0U, Name = figure.AssetPath }, new PIC.NonVisualPictureDrawingProperties()), new PIC.BlipFill(new A.Blip { Embed = relation }, new A.Stretch(new A.FillRectangle())), new PIC.ShapeProperties(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = width, Cy = height }), new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));
        return new Paragraph(drawing);
    }

    private static long EffectiveWidth(Paragraph anchor)
    {
        var body = anchor.Ancestors<Body>().FirstOrDefault(); var section = body?.Elements<SectionProperties>().LastOrDefault(); var page = section?.GetFirstChild<PageSize>(); var margins = section?.GetFirstChild<PageMargin>();
        var width = page?.Width?.Value ?? 0; var left = margins?.Left?.Value ?? 0; var right = margins?.Right?.Value ?? 0; return Math.Max(0, ((long)width - left - right) * 635L);
    }

    private static void AppendInlines(Paragraph paragraph, IReadOnlyList<YadgInline> inlines, bool italic = false, bool bold = false)
    {
        foreach (var inline in inlines) switch (inline)
        {
            case YadgText text: var run = new Run(); if (italic || bold) run.Append(new RunProperties { Italic = italic ? new Italic() : null, Bold = bold ? new Bold() : null }); run.Append(new Text(text.Value) { Space = SpaceProcessingModeValues.Preserve }); paragraph.Append(run); break;
            case YadgHardBreak: paragraph.Append(new Run(new Break())); break;
            case YadgSoftBreak: paragraph.Append(new Run(new Text(" "))); break;
            case YadgEmphasis emphasis: AppendInlines(paragraph, emphasis.Inlines, true, bold); break;
            case YadgStrong strong: AppendInlines(paragraph, strong.Inlines, italic, true); break;
        }
    }
}
