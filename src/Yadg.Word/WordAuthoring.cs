using System.Text.RegularExpressions;
using System.Security.Cryptography;
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
public sealed record CaptionPrototype(string Id, Paragraph Paragraph, int SequenceFields, int CaptionPlaceholders);
public sealed record TemplateMetadata(
    int Version,
    IReadOnlyDictionary<string, string> Styles,
    IReadOnlyDictionary<string, string> PrototypeBindings,
    IReadOnlyDictionary<string, CaptionPrototype> Prototypes,
    IReadOnlyList<Paragraph> ControlParagraphs)
{
    public static TemplateMetadata Defaults => new(0, new Dictionary<string, string>(StringComparer.Ordinal) {
        ["heading1"] = "Heading1", ["heading2"] = "Heading2", ["heading3"] = "Heading3", ["heading4"] = "Heading4", ["heading5"] = "Heading5", ["heading6"] = "Heading6",
        ["unordered"] = "ListBullet", ["ordered"] = "ListNumber", ["generatedTable"] = "TableGrid", ["caption"] = "Caption" },
        new Dictionary<string, string>(StringComparer.Ordinal), new Dictionary<string, CaptionPrototype>(StringComparer.Ordinal), Array.Empty<Paragraph>());
}
public sealed record TemplateAnalysis(IReadOnlyList<TemplatePlacement> Placements, IReadOnlyList<Diagnostic> Diagnostics, TemplateMetadata Metadata)
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

    private static TemplateAnalysis AnalyzeOpen(WordprocessingDocument document, string templatePath, YadgDocument model, TemplateMetadata? supplied = null)
    {
        var diagnostics = new List<Diagnostic>();
        var placements = new List<TemplatePlacement>();
        var main = document.MainDocumentPart;
        if (main?.Document.Body is null) return new(placements, new[] { new Diagnostic("YADG-WORD-001", "DOCX has no main document body.", true, templatePath) }, TemplateMetadata.Defaults);
        var metadata = supplied ?? ReadMetadata(main.Document.Body, templatePath, diagnostics);
        var styles = main.StyleDefinitionsPart?.Styles;
        var numbering = main.NumberingDefinitionsPart?.Numbering;
        AnalyzeParagraphs(main.Document.Body.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, diagnostics, true);
        foreach (var header in main.HeaderParts) AnalyzeParagraphs(header.Header.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, diagnostics, false);
        foreach (var footer in main.FooterParts) AnalyzeParagraphs(footer.Footer.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, diagnostics, false);
        foreach (var duplicate in placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure).GroupBy(p => (p.Kind, p.Id)).Where(g => g.Count() > 1))
            diagnostics.Add(new("YADG-PLACEMENT-002", $"Structured object '{duplicate.Key.Id}' has more than one direct placement in this template.", true, templatePath));
        ValidateReferences(document, templatePath, model, metadata, placements, diagnostics);
        return new(placements, diagnostics, metadata);
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
        var outputDiagnostics = new List<Diagnostic>();
        var metadata = ReadMetadata(document.MainDocumentPart!.Document.Body!, outputPath, outputDiagnostics);
        RemoveControlRegion(metadata);
        document.MainDocumentPart.Document.Save();
        var analysisOutput = AnalyzeOpen(document, outputPath, model, metadata);
        if (!analysisOutput.IsValid) throw new WordAuthoringException(analysisOutput.Diagnostics.First(d => d.IsError));
        var direct = analysisOutput.Placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure).Select(p => (p.Kind, p.Id)).ToHashSet();
        var targetMap = BuildTargets(model, analysisOutput, document, outputPath);
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
                elements.AddRange(CreateElements(block, placement.Anchor, document, metadata, targetMap));
            }
            foreach (var element in elements) placement.Anchor.InsertBeforeSelf(element);
            placement.Anchor.Remove();
        }
        document.MainDocumentPart!.Document.Save();
    }

    private static void AnalyzeParagraphs(IEnumerable<Paragraph> paragraphs, string templatePath, YadgDocument model, Styles? styles, Numbering? numbering, TemplateMetadata metadata, List<TemplatePlacement> placements, List<Diagnostic> diagnostics, bool mainBody)
    {
        foreach (var paragraph in paragraphs)
        {
            if (metadata.ControlParagraphs.Contains(paragraph)) continue;
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
                    if (kind == PlacementKind.Table) ValidateTable(model.FindTable(id)!, styles, metadata, templatePath, diagnostics);
                    else ValidateFigure(model.FindFigure(id)!, styles, metadata, paragraph, templatePath, diagnostics);
                    placements.Add(new(paragraph, kind, id)); continue;
                }
                if (SectionReference.TryParseTag(match.Value, out var sectionReference, out _))
                {
                    var resolved = MarkdownDocumentParser.Resolve(model, sectionReference!, templatePath);
                    if (resolved.Diagnostic is not null) { diagnostics.Add(resolved.Diagnostic); continue; }
                    ValidateBlocks(resolved.Blocks!, styles, numbering, metadata, paragraph, templatePath, diagnostics);
                    placements.Add(new(paragraph, PlacementKind.Section, sectionReference!.Id, sectionReference));
                }
                else if (match.Value.StartsWith("{{value:", StringComparison.Ordinal)) diagnostics.Add(new("YADG-VALUE-UNSUPPORTED", $"Reserved value tag '{match.Value}' is not supported in M0003.", true, templatePath));
                else diagnostics.Add(new("YADG-TAG-001", $"Malformed or unsupported tag '{match.Value}'.", true, templatePath));
            }
        }
    }

    private static void ValidateBlocks(IEnumerable<YadgBlock> blocks, Styles? styles, Numbering? numbering, TemplateMetadata metadata, Paragraph anchor, string location, List<Diagnostic> diagnostics)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case YadgHeading heading when !HasStyle(styles, StyleFor(metadata, $"heading{heading.Level}", $"Heading{heading.Level}")): diagnostics.Add(new("YADG-WORD-STYLE", $"Missing required Word style '{StyleFor(metadata, $"heading{heading.Level}", $"Heading{heading.Level}")}'.", true, location)); break;
                case YadgList list: ValidateList(list, styles, numbering, metadata, location, diagnostics); break;
                case YadgTable table: ValidateTable(table, styles, metadata, location, diagnostics); break;
                case YadgFigure figure: ValidateFigure(figure, styles, metadata, anchor, location, diagnostics); break;
            }
        }
    }

    private static void ValidateList(YadgList list, Styles? styles, Numbering? numbering, TemplateMetadata metadata, string location, List<Diagnostic> diagnostics)
    {
        var style = StyleFor(metadata, list.Ordered ? "ordered" : "unordered", list.Ordered ? "ListNumber" : "ListBullet");
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

    private static void ValidateTable(YadgTable table, Styles? styles, TemplateMetadata metadata, string location, List<Diagnostic> diagnostics)
    { if (!HasStyle(styles, StyleFor(metadata, "generatedTable", "TableGrid"))) diagnostics.Add(new("YADG-WORD-TABLE", $"Missing required Word table style '{StyleFor(metadata, "generatedTable", "TableGrid")}'.", true, location)); if (!string.IsNullOrWhiteSpace(table.Caption) && !metadata.PrototypeBindings.ContainsKey("tableCaption") && !HasStyle(styles, StyleFor(metadata, "caption", "Caption"))) diagnostics.Add(new("YADG-WORD-STYLE", "Missing required Word style for table caption.", true, location)); }

    private static void ValidateFigure(YadgFigure figure, Styles? styles, TemplateMetadata metadata, Paragraph anchor, string location, List<Diagnostic> diagnostics)
    {
        if (figure.Asset is null) { diagnostics.Add(new("YADG-FIGURE-008", $"Figure '{figure.Id}' has no validated image asset.", true, location)); return; }
        if (!string.IsNullOrEmpty(figure.AltText) && !metadata.PrototypeBindings.ContainsKey("figureCaption") && !HasStyle(styles, StyleFor(metadata, "caption", "Caption"))) diagnostics.Add(new("YADG-WORD-STYLE", "Missing required Word style for caption.", true, location));
        if (EffectiveWidth(anchor) <= 0) diagnostics.Add(new("YADG-FIGURE-009", $"Cannot determine effective text width for figure '{figure.Id}'.", true, location));
    }

    private static bool HasStyle(Styles? styles, string id) => styles?.Descendants<Style>().Any(s => s.StyleId?.Value == id) == true;
    private static string StyleFor(TemplateMetadata metadata, string role, string fallback) => metadata.Styles.TryGetValue(role, out var value) ? value : fallback;

    private static IEnumerable<OpenXmlElement> CreateElements(YadgBlock block, Paragraph anchor, WordprocessingDocument document, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets)
    {
        switch (block)
        {
            case YadgHeading heading: yield return CreateParagraph(heading, anchor, metadata, targets); break;
            case YadgParagraph paragraph: yield return CreateParagraph(paragraph, anchor, metadata, targets); break;
            case YadgList list:
                foreach (var item in list.Items) yield return CreateListParagraph(item, list.Ordered, metadata, targets); break;
            case YadgTable table:
                yield return CreateTable(table, metadata, targets);
                if (!string.IsNullOrWhiteSpace(table.Caption)) yield return CreateCaption(table.Id, table.Caption!, "tableCaption", metadata, targets);
                break;
            case YadgFigure figure:
                var imageParagraph = CreateFigureParagraph(figure, anchor, document); yield return imageParagraph;
                if (!string.IsNullOrEmpty(figure.AltText)) yield return CreateCaption(figure.Id, figure.AltText, "figureCaption", metadata, targets);
                break;
        }
    }

    private static Paragraph CreateParagraph(YadgBlock block, Paragraph anchor, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets)
    {
        var paragraph = new Paragraph();
        if (block is YadgHeading heading) paragraph.Append(new ParagraphProperties(new ParagraphStyleId { Val = StyleFor(metadata, $"heading{heading.Level}", $"Heading{heading.Level}") }));
        else if (anchor.ParagraphProperties is not null) paragraph.Append(anchor.ParagraphProperties.CloneNode(true));
        AppendInlines(paragraph, block is YadgHeading h ? new[] { new YadgText(h.Text) } : ((YadgParagraph)block).Inlines, targets);
        if (block is YadgHeading headingWithId && headingWithId.Id is not null && targets.TryGetValue(headingWithId.Id, out var headingBookmark) && headingBookmark.StartsWith("section:", StringComparison.Ordinal)) WrapParagraphBookmark(paragraph, headingBookmark[8..]);
        return paragraph;
    }

    private static Paragraph CreateListParagraph(YadgListItem item, bool ordered, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets)
    {
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = StyleFor(metadata, ordered ? "ordered" : "unordered", ordered ? "ListNumber" : "ListBullet") })); AppendInlines(paragraph, item.Inlines, targets); return paragraph;
    }

    private static Table CreateTable(YadgTable table, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets)
    {
        var result = new Table(new TableProperties(new TableStyle { Val = StyleFor(metadata, "generatedTable", "TableGrid") }, new TableLayout { Type = TableLayoutValues.Autofit }, new TableLook { FirstRow = OnOffValue.FromBoolean(true) }));
        result.Append(new TableGrid(Enumerable.Range(0, table.Header.Count).Select(_ => new GridColumn())));
        var header = new TableRow(new TableRowProperties(new TableHeader())); foreach (var cell in table.Header) header.Append(CreateCell(cell, targets)); result.Append(header);
        foreach (var row in table.Rows) { var tableRow = new TableRow(); foreach (var cell in row) tableRow.Append(CreateCell(cell, targets)); result.Append(tableRow); }
        return result;
    }

    private static TableCell CreateCell(IReadOnlyList<YadgInline> inlines, IReadOnlyDictionary<string, string> targets) { var paragraph = new Paragraph(); AppendInlines(paragraph, inlines, targets); return new TableCell(paragraph); }

    private static Paragraph CreateFigureParagraph(YadgFigure figure, Paragraph anchor, WordprocessingDocument document)
    {
        var asset = figure.Asset!; var main = document.MainDocumentPart!; var partType = asset.Format == "PNG" ? ImagePartType.Png : ImagePartType.Jpeg; var part = main.AddImagePart(partType);
        using (var stream = File.OpenRead(asset.FullPath)) part.FeedData(stream);
        var relation = main.GetIdOfPart(part); var width = Math.Min((long)asset.WidthPixels * 9525, EffectiveWidth(anchor)); var height = Math.Max(1, (long)asset.HeightPixels * 9525 * width / ((long)asset.WidthPixels * 9525));
        var drawing = new Drawing(new DW.Inline(new DW.Extent { Cx = width, Cy = height }, new DW.DocProperties { Id = 1U, Name = figure.Id }, new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }), new A.Graphic(new A.GraphicData(new PIC.Picture(new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties { Id = 0U, Name = figure.AssetPath }, new PIC.NonVisualPictureDrawingProperties()), new PIC.BlipFill(new A.Blip { Embed = relation }, new A.Stretch(new A.FillRectangle())), new PIC.ShapeProperties(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = width, Cy = height }), new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));
        return new Paragraph(drawing);
    }

    private static Paragraph CreateCaption(string targetId, string text, string binding, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets)
    {
        if (metadata.PrototypeBindings.TryGetValue(binding, out var id) && metadata.Prototypes.TryGetValue(id, out var prototype))
        {
            var clone = (Paragraph)prototype.Paragraph.CloneNode(true);
            ReplaceLogicalText(clone, "{{caption}}", text, targets);
            if (targets.TryGetValue(targetId, out var bookmark)) WrapSequenceInBookmark(clone, bookmark);
            return clone;
        }
        var plain = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = StyleFor(metadata, "caption", "Caption") }));
        AppendInlines(plain, new[] { new YadgText(text) }, targets);
        return plain;
    }

    private static long EffectiveWidth(Paragraph anchor)
    {
        var body = anchor.Ancestors<Body>().FirstOrDefault(); var section = body?.Elements<SectionProperties>().LastOrDefault(); var page = section?.GetFirstChild<PageSize>(); var margins = section?.GetFirstChild<PageMargin>();
        var width = page?.Width?.Value ?? 0; var left = margins?.Left?.Value ?? 0; var right = margins?.Right?.Value ?? 0; return Math.Max(0, ((long)width - left - right) * 635L);
    }

    private static void AppendInlines(Paragraph paragraph, IReadOnlyList<YadgInline> inlines, IReadOnlyDictionary<string, string> targets, bool italic = false, bool bold = false)
    {
        foreach (var inline in inlines) switch (inline)
        {
            case YadgText text: var run = new Run(); if (italic || bold) run.Append(new RunProperties { Italic = italic ? new Italic() : null, Bold = bold ? new Bold() : null }); run.Append(new Text(text.Value) { Space = SpaceProcessingModeValues.Preserve }); paragraph.Append(run); break;
            case YadgReference reference:
                if (targets.TryGetValue(reference.Id, out var bookmark)) paragraph.Append(CreateRefRun(bookmark));
                else paragraph.Append(new Run(new Text($"[@{reference.Id}]") { Space = SpaceProcessingModeValues.Preserve }));
                break;
            case YadgHardBreak: paragraph.Append(new Run(new Break())); break;
            case YadgSoftBreak: paragraph.Append(new Run(new Text(" "))); break;
            case YadgEmphasis emphasis: AppendInlines(paragraph, emphasis.Inlines, targets, true, bold); break;
            case YadgStrong strong: AppendInlines(paragraph, strong.Inlines, targets, italic, true); break;
        }
    }

    private static Run CreateRefRun(string target)
    {
        var paragraphNumber = target.StartsWith("section:", StringComparison.Ordinal);
        var bookmark = paragraphNumber ? target[8..] : target;
        var instruction = paragraphNumber ? $" REF {bookmark} \\p \\h " : $" REF {bookmark} \\h ";
        return new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }, new FieldCode(instruction), new FieldChar { FieldCharType = FieldCharValues.Separate }, new Text("0"), new FieldChar { FieldCharType = FieldCharValues.End });
    }

    private static TemplateMetadata ReadMetadata(Body body, string location, List<Diagnostic> diagnostics)
    {
        var paragraphs = body.Elements<Paragraph>().ToArray();
        var nonEmpty = paragraphs.FirstOrDefault(p => !string.IsNullOrWhiteSpace(LogicalText(p)));
        if (nonEmpty is null || !string.Equals(LogicalText(nonEmpty).Trim(), "{{yadg:frontmatter}}", StringComparison.Ordinal)) return TemplateMetadata.Defaults;
        var controls = new List<Paragraph> { nonEmpty };
        var lines = new List<string>(); var end = -1;
        var start = Array.IndexOf(paragraphs, nonEmpty) + 1;
        for (var i = start; i < paragraphs.Length; i++)
        {
            var text = LogicalText(paragraphs[i]); controls.Add(paragraphs[i]);
            if (text.Trim() == "{{/yadg:frontmatter}}") { end = i; break; }
            lines.Add(text);
        }
        if (end < 0) { diagnostics.Add(new("YADG-FM-001", "Front matter has no closing marker.", true, location)); return TemplateMetadata.Defaults with { ControlParagraphs = controls }; }
        var styles = new Dictionary<string, string>(TemplateMetadata.Defaults.Styles, StringComparer.Ordinal); var bindings = new Dictionary<string, string>(StringComparer.Ordinal); var state = ""; var versionSeen = false;
        foreach (var raw in lines)
        {
            var indent = raw.Length - raw.TrimStart().Length; var line = raw.Trim(); if (line.Length == 0) continue;
            var colon = line.IndexOf(':'); if (colon < 0) { diagnostics.Add(new("YADG-FM-002", $"Malformed front matter line '{raw}'.", true, location)); continue; }
            var key = line[..colon].Trim(); var value = line[(colon + 1)..].Trim().Trim('\"', '\'');
            if (key == "version") { if (!versionSeen && value != "1") diagnostics.Add(new("YADG-FM-003", "Front matter version must be 1.", true, location)); else if (versionSeen) diagnostics.Add(new("YADG-FM-006", "Duplicate front matter key 'version'.", true, location)); versionSeen = true; state = "version"; continue; }
            if (key is "styles" or "headings" or "lists" or "prototypes") { state = key; continue; }
            if (key == "generatedTable" || key == "caption") { if (indent > 2) diagnostics.Add(new("YADG-FM-004", $"Unknown front matter key '{key}'.", true, location)); else styles[key] = value; continue; }
            if (state == "headings" && Regex.IsMatch(key, "^[1-6]$")) { styles[$"heading{key}"] = value; continue; }
            if (state == "lists" && key is "unordered" or "ordered") { styles[key] = value; continue; }
            if (state == "prototypes" && key is "figureCaption" or "tableCaption") { if (!bindings.TryAdd(key, value)) diagnostics.Add(new("YADG-FM-005", $"Duplicate front matter binding '{key}'.", true, location)); continue; }
            diagnostics.Add(new("YADG-FM-006", $"Unknown front matter key '{key}'.", true, location));
        }
        if (!versionSeen) diagnostics.Add(new("YADG-FM-003", "Front matter must declare version: 1.", true, location));
        var prototypeMap = new Dictionary<string, CaptionPrototype>(StringComparer.Ordinal); var index = end + 1;
        while (index < paragraphs.Length && LogicalText(paragraphs[index]).Trim().StartsWith("{{yadg:prototype:", StringComparison.Ordinal))
        {
            var marker = LogicalText(paragraphs[index]).Trim(); var id = Regex.Match(marker, @"^\{\{yadg:prototype:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$"); controls.Add(paragraphs[index]); index++;
            if (!id.Success) { diagnostics.Add(new("YADG-PROT-001", $"Malformed prototype marker '{marker}'.", true, location)); continue; }
            var inside = new List<Paragraph>();
            while (index < paragraphs.Length && !LogicalText(paragraphs[index]).Trim().Equals($"{{{{/yadg:prototype:{id.Groups["id"].Value}}}}}", StringComparison.Ordinal)) { inside.Add(paragraphs[index]); controls.Add(paragraphs[index]); index++; }
            if (index >= paragraphs.Length) { diagnostics.Add(new("YADG-PROT-002", $"Prototype '{id.Groups["id"].Value}' has no closing marker.", true, location)); break; }
            controls.Add(paragraphs[index]); index++;
            if (inside.Count != 1) { diagnostics.Add(new("YADG-PROT-003", $"Prototype '{id.Groups["id"].Value}' must contain exactly one paragraph.", true, location)); continue; }
            var clone = (Paragraph)inside[0].CloneNode(true); var text = LogicalText(clone); var seq = clone.Descendants<SimpleField>().Count() + clone.Descendants<FieldCode>().Count(t => t.Text?.Contains("SEQ", StringComparison.OrdinalIgnoreCase) == true);
            var proto = new CaptionPrototype(id.Groups["id"].Value, clone, seq, Regex.Matches(text, Regex.Escape("{{caption}}")).Count);
            if (!prototypeMap.TryAdd(proto.Id, proto)) diagnostics.Add(new("YADG-PROT-004", $"Duplicate prototype ID '{proto.Id}'.", true, location));
            if (seq != 1 || proto.CaptionPlaceholders != 1) diagnostics.Add(new("YADG-PROT-005", $"Caption prototype '{proto.Id}' must contain exactly one SEQ field and one {{caption}} placeholder.", true, location));
            if (Regex.Matches(text, @"\{\{[^{}]+\}\}").Cast<Match>().Any(m => !string.Equals(m.Value, "{{caption}}", StringComparison.Ordinal))) diagnostics.Add(new("YADG-PROT-007", $"Caption prototype '{proto.Id}' contains another YADG placeholder.", true, location));
        }
        if (paragraphs.Skip(index).Any(p => LogicalText(p).Trim().Equals("{{yadg:frontmatter}}", StringComparison.Ordinal))) diagnostics.Add(new("YADG-FM-007", "Duplicate front matter is not supported.", true, location));
        foreach (var binding in bindings.Values) if (!prototypeMap.ContainsKey(binding)) diagnostics.Add(new("YADG-PROT-006", $"Configured prototype '{binding}' does not exist.", true, location));
        return new(1, styles, bindings, prototypeMap, controls);
    }

    private static string LogicalText(OpenXmlElement element) => string.Concat(element.Descendants<Text>().Select(t => t.Text));
    private static void RemoveControlRegion(TemplateMetadata metadata) { foreach (var paragraph in metadata.ControlParagraphs.ToArray()) paragraph.Remove(); }

    private static void ReplaceLogicalText(Paragraph paragraph, string oldText, string replacement, IReadOnlyDictionary<string, string> _)
    {
        var texts = paragraph.Descendants<Text>().ToArray(); var combined = string.Concat(texts.Select(t => t.Text)); var at = combined.IndexOf(oldText, StringComparison.Ordinal); if (at < 0) return;
        var endAt = at + oldText.Length; var offset = 0; var startIndex = 0; var endIndex = 0;
        for (var i = 0; i < texts.Length; i++) { var next = offset + texts[i].Text.Length; if (at >= offset && at < next) startIndex = i; if (endAt > offset && endAt <= next) { endIndex = i; break; } offset = next; }
        var startOffset = texts.Take(startIndex).Sum(t => t.Text.Length); var endOffset = texts.Take(endIndex).Sum(t => t.Text.Length);
        var startLocal = at - startOffset; var endLocal = endAt - endOffset;
        if (startIndex == endIndex) texts[startIndex].Text = texts[startIndex].Text[..startLocal] + replacement + texts[startIndex].Text[endLocal..];
        else
        {
            texts[startIndex].Text = texts[startIndex].Text[..startLocal] + replacement;
            for (var i = startIndex + 1; i < endIndex; i++) texts[i].Text = string.Empty;
            texts[endIndex].Text = texts[endIndex].Text[endLocal..];
        }
        foreach (var text in texts) text.Space = SpaceProcessingModeValues.Preserve;
    }

    private static void WrapSequenceInBookmark(Paragraph paragraph, string name)
    {
        var safe = Regex.Replace(name, "[^A-Za-z0-9_]", "_"); if (!char.IsLetter(safe[0])) safe = "yadg_" + safe;
        var start = paragraph.Descendants<FieldChar>().FirstOrDefault(f => f.FieldCharType?.Value == FieldCharValues.Begin);
        var end = paragraph.Descendants<FieldChar>().LastOrDefault(f => f.FieldCharType?.Value == FieldCharValues.End);
        var bookmarkId = Convert.ToInt64(Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(safe)))[..8], 16).ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (start is null || end is null)
        {
            var simple = paragraph.Descendants<SimpleField>().FirstOrDefault(); if (simple is null) return;
            simple.InsertBeforeSelf(new BookmarkStart { Id = bookmarkId, Name = safe }); simple.InsertAfterSelf(new BookmarkEnd { Id = bookmarkId }); return;
        }
        OpenXmlElement startRun = start.Ancestors<Run>().FirstOrDefault() ?? (OpenXmlElement)start;
        OpenXmlElement endRun = end.Ancestors<Run>().FirstOrDefault() ?? (OpenXmlElement)end;
        startRun.InsertBeforeSelf(new BookmarkStart { Id = bookmarkId, Name = safe }); endRun.InsertAfterSelf(new BookmarkEnd { Id = bookmarkId });
    }

    private static void WrapParagraphBookmark(Paragraph paragraph, string name)
    {
        var safe = Regex.Replace(name, "[^A-Za-z0-9_]", "_"); var bookmarkId = Convert.ToInt64(Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(safe)))[..8], 16).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var firstRun = paragraph.Elements<Run>().FirstOrDefault(); var lastRun = paragraph.Elements<Run>().LastOrDefault(); if (firstRun is null || lastRun is null) return;
        firstRun.InsertBeforeSelf(new BookmarkStart { Id = bookmarkId, Name = safe }); lastRun.InsertAfterSelf(new BookmarkEnd { Id = bookmarkId });
    }

    private static void ValidateReferences(WordprocessingDocument document, string location, YadgDocument model, TemplateMetadata metadata, IReadOnlyList<TemplatePlacement> placements, List<Diagnostic> diagnostics)
    {
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles; var numbering = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var targets = BuildTargets(model, placements, metadata, styles, numbering, diagnostics, location);
        foreach (var reference in placements.SelectMany(p => p.Kind == PlacementKind.Section ? MarkdownDocumentParser.Resolve(model, p.SectionReference!, location).Blocks ?? Array.Empty<YadgBlock>() : p.Kind == PlacementKind.Table ? new YadgBlock[] { model.FindTable(p.Id)! } : new YadgBlock[] { model.FindFigure(p.Id)! }).SelectMany(ReferencesIn))
        {
            if (model.FindObject(reference.Id) is null) { diagnostics.Add(new("YADG-REF-004", $"Unresolved semantic reference '[@{reference.Id}]'.", true, location)); continue; }
            if (!targets.ContainsKey(reference.Id)) diagnostics.Add(new("YADG-REF-005", $"Semantic reference '[@{reference.Id}]' has no uniquely rendered numbered target in this template.", true, location));
        }
    }

    private static IReadOnlyDictionary<string, string> BuildTargets(YadgDocument model, TemplateAnalysis analysis, WordprocessingDocument document, string location)
    {
        var diagnostics = new List<Diagnostic>(); var result = BuildTargets(model, analysis.Placements, analysis.Metadata, document.MainDocumentPart?.StyleDefinitionsPart?.Styles, document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering, diagnostics, location);
        if (diagnostics.Any(d => d.IsError)) throw new WordAuthoringException(diagnostics.First(d => d.IsError));
        return result;
    }

    private static Dictionary<string, string> BuildTargets(YadgDocument model, IReadOnlyList<TemplatePlacement> placements, TemplateMetadata metadata, Styles? styles, Numbering? numbering, List<Diagnostic> diagnostics, string location)
    {
        var direct = placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure).Select(p => (p.Kind, p.Id)).ToHashSet();
        var counts = new Dictionary<string, int>(StringComparer.Ordinal); var sections = new HashSet<string>(StringComparer.Ordinal);
        foreach (var placement in placements)
        {
            var blocks = placement.Kind == PlacementKind.Section ? MarkdownDocumentParser.Resolve(model, placement.SectionReference!, location).Blocks ?? Array.Empty<YadgBlock>() : placement.Kind == PlacementKind.Table ? new YadgBlock[] { model.FindTable(placement.Id)! } : new YadgBlock[] { model.FindFigure(placement.Id)! };
            foreach (var block in blocks)
            {
                if (placement.Kind == PlacementKind.Section && block is YadgHeading heading && placement.SectionReference!.Selection == SectionSelection.Section && heading.Id is not null) { counts[heading.Id] = counts.GetValueOrDefault(heading.Id) + 1; sections.Add(heading.Id); }
                if (block is YadgTable table && (!direct.Contains((PlacementKind.Table, table.Id)) || placement.Kind == PlacementKind.Table)) counts[table.Id] = counts.GetValueOrDefault(table.Id) + 1;
                if (block is YadgFigure figure && (!direct.Contains((PlacementKind.Figure, figure.Id)) || placement.Kind == PlacementKind.Figure)) counts[figure.Id] = counts.GetValueOrDefault(figure.Id) + 1;
            }
        }
        foreach (var p in placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure)) counts[p.Id] = counts.GetValueOrDefault(p.Id) + (direct.Count(x => x == (p.Kind, p.Id)) == 1 ? 0 : 0);
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (id, count) in counts)
        {
            var target = model.FindObject(id); var valid = count == 1;
            if (target is YadgFigure figure) valid &= !string.IsNullOrWhiteSpace(figure.AltText) && metadata.PrototypeBindings.ContainsKey("figureCaption") && metadata.Prototypes.ContainsKey(metadata.PrototypeBindings["figureCaption"]);
            if (target is YadgTable table) valid &= !string.IsNullOrWhiteSpace(table.Caption) && metadata.PrototypeBindings.ContainsKey("tableCaption") && metadata.Prototypes.ContainsKey(metadata.PrototypeBindings["tableCaption"]);
            if (target is YadgSection section) valid &= sections.Contains(id) && HasEffectiveNumbering(section.Heading.Level, metadata, styles, numbering);
            if (valid) result[id] = (target is YadgSection ? "section:" : "") + BookmarkName(id);
        }
        return result;
    }

    private static IEnumerable<YadgReference> ReferencesIn(YadgBlock block)
    {
        IEnumerable<YadgReference> Inline(IEnumerable<YadgInline> values) => values.SelectMany(i => i switch { YadgReference r => new[] { r }, YadgEmphasis e => Inline(e.Inlines), YadgStrong s => Inline(s.Inlines), _ => Array.Empty<YadgReference>() });
        return block switch { YadgParagraph p => Inline(p.Inlines), YadgList l => l.Items.SelectMany(i => Inline(i.Inlines)), YadgTable t => t.Header.SelectMany(Inline).Concat(t.Rows.SelectMany(r => r.SelectMany(Inline))), _ => Array.Empty<YadgReference>() };
    }

    private static bool HasEffectiveNumbering(int level, TemplateMetadata metadata, Styles? styles, Numbering? numbering)
    {
        var styleId = StyleFor(metadata, $"heading{level}", $"Heading{level}"); return styles is not null && ResolveNumberingId(styles, styleId) is int id && numbering?.Descendants<NumberingInstance>().Any(n => n.NumberID?.Value == id) == true;
    }

    private static string BookmarkName(string id)
    {
        using var sha = SHA256.Create(); var hash = Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(id)))[..10]; return "yadg_" + Regex.Replace(id, "[^A-Za-z0-9_]", "_") + "_" + hash;
    }
}
