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
public sealed record TemplatePlacement(Paragraph Anchor, PlacementKind Kind, string Id, SectionReference? SectionReference = null, int ContextLevel = 0, int RootLevel = 0);
public sealed record PreparedTableBinding(Table Table, TableRow MarkerRow, TableRow PrototypeRow, string Id);
public sealed record CaptionPrototype(string Id, Paragraph Paragraph, int SequenceFields, int CaptionPlaceholders);
public sealed record TemplateMetadata(
    int Version,
    IReadOnlyDictionary<string, string> Styles,
    IReadOnlyDictionary<string, string> PrototypeBindings,
    IReadOnlyDictionary<string, CaptionPrototype> Prototypes,
    IReadOnlyList<Paragraph> ControlParagraphs)
{
    public IReadOnlyDictionary<string, CaptionPrototype> ListItemPrototypes { get; init; } = new Dictionary<string, CaptionPrototype>(StringComparer.Ordinal);
    public static TemplateMetadata Defaults => new(0, new Dictionary<string, string>(StringComparer.Ordinal) {
        ["heading1"] = "Heading1", ["heading2"] = "Heading2", ["heading3"] = "Heading3", ["heading4"] = "Heading4", ["heading5"] = "Heading5", ["heading6"] = "Heading6", ["heading7"] = "Heading7", ["heading8"] = "Heading8", ["heading9"] = "Heading9",
        ["unordered"] = "ListBullet", ["ordered"] = "ListNumber", ["generatedTable"] = "TableGrid", ["caption"] = "Caption" },
        new Dictionary<string, string>(StringComparer.Ordinal), new Dictionary<string, CaptionPrototype>(StringComparer.Ordinal), Array.Empty<Paragraph>());
}
public sealed record TemplateAnalysis(IReadOnlyList<TemplatePlacement> Placements, IReadOnlyList<Diagnostic> Diagnostics, TemplateMetadata Metadata, IReadOnlyList<PreparedTableBinding>? PreparedBindings = null)
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
    private static readonly Regex PreparedTag = new(@"^\{\{table-rows:([A-Za-z][A-Za-z0-9_-]*)\}\}$", RegexOptions.Compiled);
    private static readonly Regex CellTag = new(@"\{\{cell\}\}", RegexOptions.Compiled);

    public static IReadOnlyList<string> InspectStyles(string templatePath)
    {
        using var document = WordprocessingDocument.Open(templatePath, false);
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles is null) return new[] { "No concrete styles are serialized." };
        var numbering = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        return styles.Elements<Style>().OrderBy(s => s.Type?.Value.ToString(), StringComparer.Ordinal).ThenBy(s => s.StyleId?.Value, StringComparer.Ordinal).Select(style =>
        {
            var type = style.Type?.Value == StyleValues.Paragraph ? "paragraph" : style.Type?.Value == StyleValues.Character ? "character" : style.Type?.Value == StyleValues.Table ? "table" : style.Type?.Value == StyleValues.Numbering ? "numbering" : "unknown";
            var id = style.StyleId?.Value ?? "(no-id)";
            var name = style.StyleName?.Val?.Value ?? id;
            var aliases = style.Aliases?.Val?.Value ?? "-";
            var visibility = new List<string>();
            if (style.GetFirstChild<Hidden>() is not null) visibility.Add("hidden"); if (style.GetFirstChild<SemiHidden>() is not null) visibility.Add("semiHidden"); if (style.GetFirstChild<UnhideWhenUsed>() is not null) visibility.Add("unhideWhenUsed"); if (style.GetFirstChild<PrimaryStyle>() is not null) visibility.Add("quickFormat");
            var number = type == "paragraph" ? (ResolveNumberingId(styles, id, numbering) is int n && numbering?.Descendants<NumberingInstance>().Any(x => x.NumberID?.Value == n) == true ? "style" : "none") : "-";
            return $"{type} \"{name}\" id={id} aliases={aliases} visibility={(visibility.Count == 0 ? "-" : string.Join(",", visibility))} numbering={number}";
        }).ToArray();
    }

    public static IReadOnlyList<string> InspectTemplate(string templatePath)
    {
        using var document = WordprocessingDocument.Open(templatePath, false);
        var body = document.MainDocumentPart?.Document.Body ?? throw new InvalidDataException("DOCX has no main document body.");
        var diagnostics = new List<Diagnostic>();
        var metadata = ReadMetadata(body, templatePath, diagnostics);
        var styles = document.MainDocumentPart!.StyleDefinitionsPart?.Styles;
        var numbering = document.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
        var lines = new List<string> { $"Template: {Path.GetFileName(templatePath)}", $"frontmatter: {(metadata.Version == 1 ? "present (version 1)" : "absent; compatibility defaults")}" };
        var ordinals = new Dictionary<Paragraph, int>();
        var paragraphs = document.MainDocumentPart.Document.Body!.Descendants<Paragraph>().ToArray();
        for (var i = 0; i < paragraphs.Length; i++) ordinals[paragraphs[i]] = i + 1;
        foreach (var paragraph in paragraphs)
        {
            var text = ParagraphText(paragraph);
            var match = Regex.Match(text, @"\{\{(content|section|table|figure|table-rows):([^{}]+)\}\}");
            if (match.Success) lines.Add($"placement {match.Groups[1].Value}:{match.Groups[2].Value} | {HumanLocation(paragraph, templatePath, "body", ordinals[paragraph])}");
        }
        foreach (var table in body.Descendants<Table>().Where(t => t.Ancestors<Table>().FirstOrDefault() is null))
        {
            var rows = table.Elements<TableRow>().ToArray();
            for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                var marker = Regex.Match(LogicalText(rows[rowIndex]).Trim(), @"^\{\{table-rows:([A-Za-z][A-Za-z0-9_-]*)\}\}$");
                if (!marker.Success) continue;
                var markerParagraph = rows[rowIndex].Descendants<Paragraph>().FirstOrDefault();
                var locator = markerParagraph is not null && ordinals.TryGetValue(markerParagraph, out var ordinal)
                    ? HumanLocation(markerParagraph, templatePath, "prepared table", ordinal)
                    : $"{templatePath}\nprepared table > row {rowIndex + 1}\nnear: \"{BoundExcerpt(LogicalText(rows[rowIndex]))}\"";
                var prototypeRow = rowIndex + 2 <= rows.Length ? (rowIndex + 2).ToString() : "missing";
                lines.Add($"prepared-table {marker.Groups[1].Value} marker-row={rowIndex + 1} prototype-row={prototypeRow} | {locator}");
            }
        }
        foreach (var prototype in metadata.Prototypes.OrderBy(p => p.Key, StringComparer.Ordinal))
        {
            var role = metadata.PrototypeBindings.FirstOrDefault(pair => pair.Value == prototype.Key).Key ?? "unbound";
            lines.Add($"prototype {prototype.Key} role={role} | near: {BoundExcerpt(ParagraphText(prototype.Value.Paragraph))}");
        }
        var roles = metadata.Styles.OrderBy(p => p.Key, StringComparer.Ordinal).ToArray();
        foreach (var (role, selector) in roles)
        {
            var expected = role == "generatedTable" ? StyleValues.Table : role == "codeInline" ? StyleValues.Character : StyleValues.Paragraph;
            var all = styles?.Elements<Style>().ToArray() ?? Array.Empty<Style>();
            var match = all.FirstOrDefault(s => string.Equals(s.StyleId?.Value, selector, StringComparison.Ordinal));
            if (match is null)
            {
                var candidates = all.Where(s => string.Equals(s.StyleName?.Val?.Value, selector, StringComparison.OrdinalIgnoreCase) || (s.Aliases?.Val?.Value ?? "").Split(',').Any(a => string.Equals(a.Trim(), selector, StringComparison.OrdinalIgnoreCase))).ToArray();
                if (candidates.Length == 1) match = candidates[0];
            }
            var hasListNumbering = role is "unordered" or "ordered" ? match is not null && ResolveNumberingIdForRole(all, match.StyleId?.Value, numbering, role == "ordered" ? NumberFormatValues.Decimal : NumberFormatValues.Bullet) : true;
            var roleResolved = match is not null && match.Type?.Value == expected && hasListNumbering;
            var resolved = roleResolved ? $"{match!.StyleName?.Val?.Value ?? match.StyleId?.Value} (id={match.StyleId?.Value}, {StyleTypeName(match.Type?.Value)})" : "unresolved";
            var candidate = FindRoleCandidate(all, role, expected, numbering);
            var fallback = !roleResolved ? $"; yolo candidate: {DescribeCandidate(candidate) ?? (role is "unordered" or "ordered" ? $"builtin:{(role == "ordered" ? "ordered" : "unordered")}-list-v1" : "builtin:plain-v1")}" : "";
            lines.Add($"role {role}: configured={selector}; strict={resolved}{fallback}");
        }
        foreach (var diagnostic in diagnostics) lines.Add($"warning {diagnostic.Code}: {diagnostic.Message}");
        return lines;
    }

    private static string? DescribeCandidate(Style? style) => style is null ? null : $"{style.StyleName?.Val?.Value ?? style.StyleId?.Value} (id={style.StyleId?.Value})";

    public static TemplateAnalysis Analyze(string templatePath, YadgDocument model, IReadOnlyDictionary<string, string>? values = null, bool yolo = false)
    {
        using var document = WordprocessingDocument.Open(templatePath, false);
        return AnalyzeOpen(document, templatePath, model, values ?? new Dictionary<string, string>(StringComparer.Ordinal), yolo: yolo);
    }

    private static TemplateAnalysis AnalyzeOpen(WordprocessingDocument document, string templatePath, YadgDocument model, IReadOnlyDictionary<string, string> values, TemplateMetadata? supplied = null, bool yolo = false)
    {
        var diagnostics = new List<Diagnostic>();
        var placements = new List<TemplatePlacement>();
        var prepared = new List<PreparedTableBinding>();
        var main = document.MainDocumentPart;
        if (main?.Document.Body is null) return new(placements, new[] { new Diagnostic("YADG-WORD-001", "DOCX has no main document body.", true, templatePath) }, TemplateMetadata.Defaults);
        var metadata = supplied ?? ReadMetadata(main.Document.Body, templatePath, diagnostics);
        var styles = main.StyleDefinitionsPart?.Styles;
        var numbering = main.NumberingDefinitionsPart?.Numbering;
        if (yolo) metadata = RecoverListPrototypes(metadata, styles, numbering, diagnostics, templatePath);
        metadata = ResolveSelectors(metadata, styles, numbering, diagnostics, templatePath, yolo);
        AnalyzePreparedTables(main.Document.Body, templatePath, model, prepared, diagnostics);
        AnalyzeParagraphs(main.Document.Body.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, prepared, diagnostics, values, "body", main.Document.Body, yolo);
        foreach (var header in main.HeaderParts) AnalyzeParagraphs(header.Header.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, prepared, diagnostics, values, "header", null, yolo);
        foreach (var footer in main.FooterParts) AnalyzeParagraphs(footer.Footer.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, prepared, diagnostics, values, "footer", null, yolo);
        if (main.FootnotesPart?.Footnotes is not null) AnalyzeParagraphs(main.FootnotesPart.Footnotes.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, prepared, diagnostics, values, "footnote", null, yolo);
        if (main.EndnotesPart?.Endnotes is not null) AnalyzeParagraphs(main.EndnotesPart.Endnotes.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, prepared, diagnostics, values, "endnote", null, yolo);
        if (main.WordprocessingCommentsPart?.Comments is not null) AnalyzeParagraphs(main.WordprocessingCommentsPart.Comments.Descendants<Paragraph>(), templatePath, model, styles, numbering, metadata, placements, prepared, diagnostics, values, "comment", null, yolo);
        metadata = ResolveUsedCodeInlineStyle(metadata, model, placements, prepared, styles, diagnostics, templatePath, yolo);
        foreach (var duplicate in placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure).GroupBy(p => (p.Kind, p.Id)).Where(g => g.Count() > 1))
            diagnostics.Add(new("YADG-PLACEMENT-002", $"Structured object '{duplicate.Key.Id}' has more than one direct placement in this template.", true, templatePath));
        foreach (var duplicate in prepared.GroupBy(p => p.Id).Where(g => g.Count() > 1))
            diagnostics.Add(new("YADG-PLACEMENT-002", $"Structured object '{duplicate.Key}' has more than one prepared placement in this template.", true, templatePath));
        foreach (var conflict in prepared.Select(p => p.Id).Intersect(placements.Where(p => p.Kind == PlacementKind.Table).Select(p => p.Id), StringComparer.Ordinal))
            diagnostics.Add(new("YADG-PLACEMENT-003", $"Table '{conflict}' cannot have both prepared and generated direct placement.", true, templatePath));
        ValidateReferences(document, templatePath, model, metadata, placements, prepared, diagnostics);
        if (yolo)
        {
            for (var i = diagnostics.Count - 1; i >= 0; i--)
            {
                var d = diagnostics[i];
                if (d.Code is "YADG-VALUE-002" or "YADG-VALUE-UNSUPPORTED" or "YADG-REF-004" or "YADG-REF-005")
                    diagnostics[i] = d with { Code = "YADG-YOLO-UNRESOLVED-001", IsError = false, IsDegradation = true, Message = d.Message + "; fallback: preserve the unresolved source token visibly." };
                else if (d.Code is "YADG-PROT-005" or "YADG-PROT-003" or "YADG-PROT-006")
                    diagnostics[i] = d with { Code = "YADG-YOLO-PROTOTYPE-001", IsError = false, IsDegradation = true, Message = d.Message + "; fallback: resolve presentation through a compatible resource or built-in example." };
            }
            var usedRoles = new HashSet<string>(StringComparer.Ordinal);
            foreach (var placement in placements)
            {
                var blocks = placement.Kind == PlacementKind.Section ? MarkdownDocumentParser.Resolve(model, placement.SectionReference!, templatePath).Blocks ?? Array.Empty<YadgBlock>() : placement.Kind == PlacementKind.Table ? new YadgBlock[] { model.FindTable(placement.Id)! } : new YadgBlock[] { model.FindFigure(placement.Id)! };
                foreach (var block in blocks)
                {
                    if (block is YadgHeading h) usedRoles.Add($"heading{EffectiveHeadingLevel(h.Level, placement.ContextLevel, placement.SectionReference?.Selection ?? SectionSelection.Section, placement.RootLevel)}");
                    if (block is YadgList l) usedRoles.Add(l.Ordered ? "ordered" : "unordered");
                    if (block is YadgTable t) { usedRoles.Add("generatedTable"); if (!string.IsNullOrWhiteSpace(t.Caption)) usedRoles.Add("caption"); }
                    if (block is YadgFigure f && !string.IsNullOrWhiteSpace(f.AltText)) usedRoles.Add("caption");
                    if (HasInlineCode(block)) usedRoles.Add("codeInline");
                }
            }
            if (prepared.Count > 0) { usedRoles.Add("generatedTable"); if (prepared.Any(p => !string.IsNullOrWhiteSpace(model.FindTable(p.Id)?.Caption))) usedRoles.Add("caption"); }
            diagnostics.RemoveAll(d => d.IsDegradation && !IsUsedRoleDegradation(d, usedRoles));
        }
        return new(placements, diagnostics, metadata, prepared);
    }

    private static bool HasInlineCode(YadgBlock block) => block switch
    {
        YadgParagraph p => p.Inlines.Any(HasCode),
        YadgList l => l.Items.Any(i => i.Inlines.Any(HasCode)),
        YadgTable t => t.Header.Concat(t.Rows.SelectMany(r => r)).Any(c => c.Any(HasCode)),
        _ => false
    };
    private static bool HasCode(YadgInline i) => i is YadgCode || i is YadgEmphasis e && e.Inlines.Any(HasCode) || i is YadgStrong s && s.Inlines.Any(HasCode);
    private static bool IsUsedRoleDegradation(Diagnostic diagnostic, HashSet<string> used)
    {
        if (diagnostic.Code is "YADG-YOLO-LIST-001" or "YADG-YOLO-LIST-002" or "YADG-YOLO-LIST-003") return used.Contains(diagnostic.Message.Contains("unordered", StringComparison.OrdinalIgnoreCase) ? "unordered" : "ordered");
        if (diagnostic.Code == "YADG-YOLO-STYLE-001")
        {
            var role = Regex.Match(diagnostic.Message, @"role '([^']+)'");
            if (!role.Success) role = Regex.Match(diagnostic.Message, @"role '?(heading[1-9])'?");
            return role.Success ? used.Contains(role.Groups[1].Value) : true;
        }
        if (diagnostic.Code == "YADG-YOLO-TABLE-001") return used.Contains("generatedTable");
        if (diagnostic.Code == "YADG-YOLO-CAPTION-001") return used.Contains("caption");
        if (diagnostic.Code == "YADG-YOLO-PROTOTYPE-001") return diagnostic.Message.Contains("unordered", StringComparison.OrdinalIgnoreCase) ? used.Contains("unordered") : diagnostic.Message.Contains("ordered", StringComparison.OrdinalIgnoreCase) ? used.Contains("ordered") : true;
        return true;
    }

    private static TemplateMetadata RecoverListPrototypes(TemplateMetadata metadata, Styles? styles, Numbering? numbering, List<Diagnostic> diagnostics, string location)
    {
        var bindings = new Dictionary<string, string>(metadata.PrototypeBindings, StringComparer.Ordinal);
        var prototypes = new Dictionary<string, CaptionPrototype>(metadata.ListItemPrototypes, StringComparer.Ordinal);
        var allPrototypes = new Dictionary<string, CaptionPrototype>(metadata.Prototypes, StringComparer.Ordinal);
        foreach (var role in new[] { "unorderedListItem", "orderedListItem" })
        {
            if (!bindings.TryGetValue(role, out var id)) continue;
            var valid = prototypes.TryGetValue(id, out var prototype) && Regex.Matches(ParagraphText(prototype!.Paragraph), Regex.Escape("{{item}}" )).Count == 1 && ResolveParagraphNumberingId(prototype.Paragraph, styles, numbering) is int numId && numbering?.Elements<NumberingInstance>().Any(n => n.NumberID?.Value == numId) == true;
            if (valid) continue;
            bindings.Remove(role); prototypes.Remove(id);
            diagnostics.Add(new("YADG-YOLO-PROTOTYPE-001", $"Configured {role} prototype '{id}' is missing or unusable; fallback: deterministic list resource resolution.", false, location) { IsDegradation = true });
        }
        foreach (var role in new[] { "figureCaption", "tableCaption" })
        {
            if (!bindings.TryGetValue(role, out var id)) continue;
            if (allPrototypes.TryGetValue(id, out var prototype) && prototype.SequenceFields == 1 && prototype.CaptionPlaceholders == 1) continue;
            bindings.Remove(role); allPrototypes.Remove(id);
            diagnostics.Add(new("YADG-YOLO-PROTOTYPE-001", $"Configured {role} prototype '{id}' is missing or unusable; fallback: plain caption without invented numbering.", false, location) { IsDegradation = true });
        }
        return metadata with { PrototypeBindings = bindings, ListItemPrototypes = prototypes, Prototypes = allPrototypes };
    }

    public static void Author(string templatePath, string outputPath, YadgDocument model, IReadOnlyDictionary<string, string>? values = null, bool yolo = false)
    {
        values ??= new Dictionary<string, string>(StringComparer.Ordinal);
        using (var analysisDocument = WordprocessingDocument.Open(templatePath, false))
        {
            var analysis = AnalyzeOpen(analysisDocument, templatePath, model, values, yolo: yolo);
            if (!analysis.IsValid) throw new WordAuthoringException(analysis.Diagnostics.First(d => d.IsError));
        }
        File.Copy(templatePath, outputPath, true);
        using var document = WordprocessingDocument.Open(outputPath, true);
        var outputDiagnostics = new List<Diagnostic>();
        var metadata = ReadMetadata(document.MainDocumentPart!.Document.Body!, outputPath, outputDiagnostics);
        RemoveControlRegion(metadata);
        document.MainDocumentPart.Document.Save();
        var analysisOutput = AnalyzeOpen(document, outputPath, model, values, metadata, yolo);
        if (!analysisOutput.IsValid) throw new WordAuthoringException(analysisOutput.Diagnostics.First(d => d.IsError));
        metadata = analysisOutput.Metadata;
        var direct = analysisOutput.Placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure).Select(p => (p.Kind, p.Id)).ToHashSet();
        foreach (var binding in analysisOutput.PreparedBindings ?? Array.Empty<PreparedTableBinding>()) direct.Add((PlacementKind.Table, binding.Id));
        var targetMap = BuildTargets(model, analysisOutput, document, outputPath);
        foreach (var binding in analysisOutput.PreparedBindings ?? Array.Empty<PreparedTableBinding>())
        {
            var table = model.FindTable(binding.Id)!;
            PopulatePreparedTable(binding, table, targetMap, metadata);
            if (!string.IsNullOrWhiteSpace(table.Caption))
                binding.Table.InsertAfterSelf(CreateCaption(table.Id, table.Caption!, "tableCaption", metadata, targetMap));
        }
        foreach (var placement in analysisOutput.Placements.Reverse())
        {
            var blocks = placement.Kind == PlacementKind.Section
                ? RebaseBlocks(MarkdownDocumentParser.Resolve(model, placement.SectionReference!, outputPath).Blocks!, placement)
                : new[] { placement.Kind == PlacementKind.Table ? (YadgBlock)model.FindTable(placement.Id)! : model.FindFigure(placement.Id)! };
            var elements = new List<OpenXmlElement>();
            foreach (var block in blocks)
            {
                if (placement.Kind == PlacementKind.Section && block is YadgTable table && direct.Contains((PlacementKind.Table, table.Id))) continue;
                if (placement.Kind == PlacementKind.Section && block is YadgFigure figure && direct.Contains((PlacementKind.Figure, figure.Id))) continue;
                elements.AddRange(CreateElements(block, placement.Anchor, document, metadata, targetMap, yolo));
            }
            foreach (var element in elements) placement.Anchor.InsertBeforeSelf(element);
            placement.Anchor.Remove();
        }
        ReplaceValues(document, values);
        document.MainDocumentPart!.Document.Save();
    }

    private static void AnalyzeParagraphs(IEnumerable<Paragraph> paragraphs, string templatePath, YadgDocument model, Styles? styles, Numbering? numbering, TemplateMetadata metadata, List<TemplatePlacement> placements, List<PreparedTableBinding> prepared, List<Diagnostic> diagnostics, IReadOnlyDictionary<string, string> values, string location, Body? body = null, bool yolo = false)
    {
        var paragraphIndex = 0;
        foreach (var paragraph in paragraphs)
        {
            paragraphIndex++;
            var structuralLocation = HumanLocation(paragraph, templatePath, location, paragraphIndex);
            if (metadata.ControlParagraphs.Contains(paragraph)) continue;
            if (prepared.Any(p => paragraph.Ancestors<TableRow>().Contains(p.MarkerRow) || paragraph.Ancestors<TableRow>().Contains(p.PrototypeRow))) continue;
            var logical = ParagraphText(paragraph);
            var matches = TagLike.Matches(logical).Cast<Match>().ToArray();
            if (matches.Length == 0) continue;
            var valueMatches = matches.Where(m => m.Value.StartsWith("{{value:", StringComparison.Ordinal)).ToArray();
            var valueLocationAllowed = IsSupportedValueLocation(location);
            if (valueMatches.Length > 0)
            {
                if (!valueLocationAllowed) diagnostics.Add(new("YADG-VALUE-LOCATION", "Workspace value tags are not supported in this Word location.", true, structuralLocation));
                foreach (var match in valueMatches)
                {
                    var id = Regex.Match(match.Value, @"^\{\{value:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$");
                    if (!id.Success) diagnostics.Add(new("YADG-VALUE-001", $"Malformed workspace value tag '{match.Value}'.", true, structuralLocation));
                    else if (!values.ContainsKey(id.Groups["id"].Value))
                    {
                        diagnostics.Add(new("YADG-VALUE-002", $"Workspace value '{id.Groups["id"].Value}' is not defined.", true, structuralLocation));
                        diagnostics.Add(new("YADG-VALUE-UNSUPPORTED", $"Workspace value tag '{match.Value}' cannot be resolved.", true, structuralLocation));
                    }
                }
            }
            var nonValueMatches = matches.Where(m => !m.Value.StartsWith("{{value:", StringComparison.Ordinal)).ToArray();
            if (valueMatches.Length > 0 && nonValueMatches.Length == 0) continue;
            if (location != "body" || paragraph.Parent is not Body)
            { diagnostics.Add(new("YADG-WORD-LOCATION", "Non-value YADG tags are supported only as standalone paragraphs in the main document body.", true, structuralLocation)); continue; }
            foreach (var match in matches)
            {
                if (match.Value.StartsWith("{{value:", StringComparison.Ordinal)) continue;
                if (!string.Equals(logical.Trim(), match.Value, StringComparison.Ordinal)) { diagnostics.Add(new("YADG-WORD-BLOCK", $"Block tag '{match.Value}' must be the only non-whitespace paragraph content.", true, structuralLocation)); continue; }
                var direct = DirectTag.Match(match.Value);
                if (direct.Success)
                {
                    var kind = direct.Groups[1].Value == "table" ? PlacementKind.Table : PlacementKind.Figure; var id = direct.Groups[2].Value;
                    if (kind == PlacementKind.Table && model.FindTable(id) is null) { diagnostics.Add(new("YADG-REF-003", $"Tag '{match.Value}' does not resolve to a table.", true, structuralLocation)); continue; }
                    if (kind == PlacementKind.Figure && model.FindFigure(id) is null) { diagnostics.Add(new("YADG-REF-003", $"Tag '{match.Value}' does not resolve to a figure.", true, structuralLocation)); continue; }
                    if (kind == PlacementKind.Table) ValidateTable(model.FindTable(id)!, styles, metadata, structuralLocation, diagnostics);
                    else ValidateFigure(model.FindFigure(id)!, styles, metadata, paragraph, structuralLocation, diagnostics, yolo);
                    placements.Add(new(paragraph, kind, id)); continue;
                }
                if (SectionReference.TryParseTag(match.Value, out var sectionReference, out _))
                {
                    var resolved = MarkdownDocumentParser.Resolve(model, sectionReference!, templatePath);
                    if (resolved.Diagnostic is not null) { diagnostics.Add(resolved.Diagnostic with { Location = structuralLocation }); continue; }
                    var context = body is not null && paragraph.Parent is Body ? ResolveTemplateContext(body, paragraph, styles) : 0;
                    var rootLevel = model.FindSection(sectionReference!.Id)!.Heading.Level;
                    ValidateBlocks(resolved.Blocks!, styles, numbering, metadata, paragraph, structuralLocation, diagnostics, context, sectionReference.Selection, rootLevel, yolo);
                    placements.Add(new(paragraph, PlacementKind.Section, sectionReference.Id, sectionReference, context, rootLevel));
                }
                else diagnostics.Add(new("YADG-TAG-001", $"Malformed or unsupported tag '{match.Value}'.", true, structuralLocation));
            }
        }
    }

    private static string HumanLocation(Paragraph paragraph, string templatePath, string story, int ordinal)
    {
        var near = BoundExcerpt(ParagraphText(paragraph));
        var context = "";
        if (story == "body" && paragraph.Parent is Body body)
        {
            var headings = new List<string>();
            foreach (var prior in body.Elements<Paragraph>().TakeWhile(p => p != paragraph))
            {
                var style = prior.ParagraphProperties?.GetFirstChild<ParagraphStyleId>()?.Val?.Value ?? "";
                if (!Regex.IsMatch(style, "^Heading[1-9]$", RegexOptions.IgnoreCase)) continue;
                var text = ParagraphText(prior); if (string.IsNullOrWhiteSpace(text)) continue;
                var level = int.Parse(style[^1..]);
                while (headings.Count >= level) headings.RemoveAt(headings.Count - 1);
                headings.Add($"\"{BoundExcerpt(text, 80)}\"");
            }
            if (headings.Count > 0) context = " > " + string.Join(" > ", headings);
        }
        return $"{templatePath}\n{story}{context} > paragraph {ordinal}\nnear: \"{near}\"";
    }

    private static string BoundExcerpt(string text, int limit = 120)
    {
        var normalized = Regex.Replace(new string(text.Where(c => !char.IsControl(c)).ToArray()), @"\s+", " ").Trim();
        return normalized.Length <= limit ? normalized : normalized[..limit] + "…";
    }

    private static void AnalyzePreparedTables(Body body, string location, YadgDocument model, List<PreparedTableBinding> bindings, List<Diagnostic> diagnostics)
    {
        foreach (var table in body.Descendants<Table>().Where(t => t.Ancestors<Table>().FirstOrDefault() is null))
        {
            var rows = table.Elements<TableRow>().ToArray();
            for (var i = 0; i < rows.Length; i++)
            {
                var markerText = LogicalText(rows[i]);
                var match = PreparedTag.Match(markerText.Trim());
                if (!match.Success)
                {
                    if (TagLike.Matches(markerText).Cast<Match>().Any(m => m.Value.StartsWith("{{table-rows:", StringComparison.Ordinal)))
                        diagnostics.Add(new("YADG-PREPARED-001", "Prepared marker row must contain only one table-rows marker and whitespace.", true, location));
                    continue;
                }
                if (!string.Equals(markerText.Trim(), match.Value, StringComparison.Ordinal) || rows[i].Elements<TableCell>().SelectMany(c => c.Descendants<Text>()).Any(t => t.Text?.Contains("{{table-rows:", StringComparison.Ordinal) == true && !markerText.Trim().Equals(match.Value, StringComparison.Ordinal)))
                    diagnostics.Add(new("YADG-PREPARED-001", "Prepared marker row must contain only one table-rows marker and whitespace.", true, location));
                var id = match.Groups[1].Value;
                var semantic = model.FindTable(id);
                if (semantic is null) { diagnostics.Add(new("YADG-REF-003", $"Prepared marker '{match.Value}' does not resolve to a table.", true, location)); continue; }
                if (i + 1 >= rows.Length) { diagnostics.Add(new("YADG-PREPARED-002", $"Prepared table '{id}' has no prototype row immediately after its marker.", true, location)); continue; }
                var prototype = rows[i + 1];
                ValidatePrototype(prototype, semantic, id, location, diagnostics);
                bindings.Add(new(table, rows[i], prototype, id));
            }
        }
    }

    private static bool IsSupportedValueLocation(string location)
        => location is "body" or "header" or "footer" or "footnote" or "endnote" or "comment";

    private static string ParagraphText(Paragraph paragraph)
        => string.Concat(paragraph.Descendants<Text>().Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph).Select(t => t.Text));

    private static int ResolveTemplateContext(Body body, Paragraph placement, Styles? styles)
    {
        foreach (var paragraph in body.Elements<Paragraph>().TakeWhile(p => !ReferenceEquals(p, placement)).Reverse())
        {
            var level = EffectiveOutlineLevel(paragraph, styles);
            if (level is >= 1 and <= 9) return level.Value;
        }
        return 0;
    }

    private static int? EffectiveOutlineLevel(Paragraph paragraph, Styles? styles)
    {
        var direct = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (direct is not null) return direct + 1;
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (styles is not null && !string.IsNullOrEmpty(styleId) && seen.Add(styleId))
        {
            var style = styles.Descendants<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style is null) break;
            var outline = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            if (outline is not null) return outline + 1;
            styleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private static void ValidatePrototype(TableRow prototype, YadgTable semantic, string id, string location, List<Diagnostic> diagnostics)
    {
        var cells = prototype.Elements<TableCell>().ToArray();
        if (cells.Length != semantic.Header.Count)
            diagnostics.Add(new("YADG-PREPARED-003", $"Prepared table '{id}' prototype has {cells.Length} cells but the semantic table has {semantic.Header.Count} columns.", true, location));
        if (prototype.Descendants<GridSpan>().Any(g => (g.Val?.Value ?? 1) > 1)) diagnostics.Add(new("YADG-PREPARED-004", $"Prepared table '{id}' prototype cannot contain horizontal grid spans.", true, location));
        if (prototype.Descendants<VerticalMerge>().Any()) diagnostics.Add(new("YADG-PREPARED-005", $"Prepared table '{id}' prototype cannot contain vertical merges.", true, location));
        foreach (var cell in cells)
        {
            var paragraphCount = cell.Elements<Paragraph>().Count();
            if (paragraphCount != 1 || cell.Elements<OpenXmlElement>().Any(e => e is not Paragraph && e is not TableCellProperties))
                diagnostics.Add(new("YADG-PREPARED-006", $"Each prepared table '{id}' prototype cell must contain exactly one ordinary paragraph.", true, location));
            if (cell.Descendants<Table>().Any() || cell.Descendants<Drawing>().Any() || cell.Descendants<SdtElement>().Any() || cell.Descendants<BookmarkStart>().Any() || cell.Descendants<BookmarkEnd>().Any() || cell.Descendants<SimpleField>().Any() || cell.Descendants<FieldChar>().Any() || cell.Descendants<FieldCode>().Any())
                diagnostics.Add(new("YADG-PREPARED-007", $"Prepared table '{id}' prototype contains a forbidden nested structure, field, bookmark, drawing, or content control.", true, location));
            var text = LogicalText(cell);
            if (CellTag.Matches(text).Count != 1)
                diagnostics.Add(new("YADG-PREPARED-008", $"Each prepared table '{id}' prototype cell must contain exactly one logical {{cell}} placeholder.", true, location));
            foreach (var tag in TagLike.Matches(text).Cast<Match>())
                if (!string.Equals(tag.Value, "{{cell}}", StringComparison.Ordinal)) diagnostics.Add(new("YADG-PREPARED-009", $"Prepared table '{id}' prototype contains forbidden YADG tag '{tag.Value}'.", true, location));
        }
    }

    private static void ValidateBlocks(IEnumerable<YadgBlock> blocks, Styles? styles, Numbering? numbering, TemplateMetadata metadata, Paragraph anchor, string location, List<Diagnostic> diagnostics, int contextLevel = 0, SectionSelection selection = SectionSelection.Section, int rootLevel = 0, bool yolo = false)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case YadgHeading heading:
                    var effective = EffectiveHeadingLevel(heading.Level, contextLevel, selection, rootLevel);
                    if (effective > 9) diagnostics.Add(new("YADG-WORD-HEADING-OVERFLOW", $"Markdown heading level {heading.Level} rebases to unsupported effective Word heading level {effective}.", true, location));
                    else if (StyleFor(metadata, $"heading{effective}", $"Heading{effective}").StartsWith("builtin:", StringComparison.Ordinal)) { }
                    else if (!HasStyle(styles, StyleFor(metadata, $"heading{effective}", $"Heading{effective}")) && !yolo) diagnostics.Add(new("YADG-WORD-STYLE", $"Missing required Word style '{StyleFor(metadata, $"heading{effective}", $"Heading{effective}")}'; run yadg inspect styles to discover available styles.", true, location));
                    else if (!HasStyle(styles, StyleFor(metadata, $"heading{effective}", $"Heading{effective}")) && yolo) diagnostics.Add(new("YADG-YOLO-HEADING-001", $"Heading style for effective level {effective} is unavailable; fallback: plain paragraph with outline level {effective}.", false, location) { IsDegradation = true });
                    break;
                case YadgList list: ValidateList(list, styles, numbering, metadata, location, diagnostics, yolo); break;
                case YadgTable table: ValidateTable(table, styles, metadata, location, diagnostics); break;
                case YadgFigure figure: ValidateFigure(figure, styles, metadata, anchor, location, diagnostics, yolo); break;
            }
        }
    }

    private static TemplateMetadata ResolveUsedCodeInlineStyle(TemplateMetadata metadata, YadgDocument model, IReadOnlyList<TemplatePlacement> placements, IReadOnlyList<PreparedTableBinding> prepared, Styles? styles, List<Diagnostic> diagnostics, string location, bool yolo = false)
    {
        IEnumerable<YadgBlock> selectedBlocks()
        {
            foreach (var placement in placements)
            {
                if (placement.Kind == PlacementKind.Section)
                {
                    var resolved = MarkdownDocumentParser.Resolve(model, placement.SectionReference!, location);
                    if (resolved.Blocks is not null) foreach (var block in resolved.Blocks) yield return block;
                }
                else if (placement.Kind == PlacementKind.Table && model.FindTable(placement.Id) is { } table) yield return table;
                else if (placement.Kind == PlacementKind.Figure && model.FindFigure(placement.Id) is { } figure) yield return figure;
            }
            foreach (var binding in prepared) if (model.FindTable(binding.Id) is { } table) yield return table;
        }
        IEnumerable<YadgInline> allInlines(YadgBlock block) => block switch
        {
            YadgParagraph paragraph => paragraph.Inlines,
            YadgList list => list.Items.SelectMany(item => item.Inlines),
            YadgTable table => table.Header.Concat(table.Rows.SelectMany(row => row)).SelectMany(cell => cell),
            _ => Array.Empty<YadgInline>()
        };
        IEnumerable<YadgInline> descendants(IEnumerable<YadgInline> inlines) => inlines.SelectMany(inline => inline switch { YadgEmphasis e => new[] { inline }.Concat(descendants(e.Inlines)), YadgStrong s => new[] { inline }.Concat(descendants(s.Inlines)), _ => new[] { inline } });
        if (!selectedBlocks().SelectMany(allInlines).SelectMany(inline => descendants(new[] { inline })).Any(inline => inline is YadgCode)) return metadata;
        if (!metadata.Styles.TryGetValue("codeInline", out var selector))
        {
            if (yolo) { diagnostics.Add(new("YADG-YOLO-CODE-001", "Styled inline code has no configured character style; fallback: preserve code text as plain text.", false, location) { IsDegradation = true }); return metadata with { Styles = new Dictionary<string, string>(metadata.Styles, StringComparer.Ordinal) { ["codeInline"] = "builtin:plain-v1" } }; }
            diagnostics.Add(new("YADG-WORD-STYLE", "Styled inline code requires a character style role 'codeInline'. Configure it in template front matter and inspect available styles with yadg inspect styles.", true, location));
            return metadata;
        }
        if (styles is null) { diagnostics.Add(new("YADG-WORD-STYLE", $"Style selector '{selector}' cannot resolve because the template has no style definitions. Run yadg inspect styles.", true, location)); return metadata; }
        var all = styles.Elements<Style>().ToArray();
        var selected = all.FirstOrDefault(s => string.Equals(s.StyleId?.Value, selector, StringComparison.Ordinal));
        if (selected is null)
        {
            var candidates = all.Where(s => string.Equals(s.StyleName?.Val?.Value, selector, StringComparison.OrdinalIgnoreCase) || (s.Aliases?.Val?.Value ?? "").Split(',').Any(a => string.Equals(a.Trim(), selector, StringComparison.OrdinalIgnoreCase))).ToArray();
            if (candidates.Length > 1) { diagnostics.Add(new("YADG-WORD-STYLE", $"Style selector '{selector}' is ambiguous: {string.Join(", ", candidates.Select(s => $"{s.StyleName?.Val?.Value ?? s.StyleId?.Value} (id={s.StyleId?.Value})"))}. Run yadg inspect styles.", true, location)); return metadata; }
            if (candidates.Length == 1) selected = candidates[0];
        }
        if (selected is null && yolo)
        {
            var related = FindRoleCandidate(styles?.Elements<Style>().ToArray() ?? Array.Empty<Style>(), "codeInline", StyleValues.Character, null);
            if (related is not null) { diagnostics.Add(new("YADG-YOLO-CODE-001", $"Configured inline-code style '{selector}' is unavailable; fallback: related character style '{related.StyleName?.Val?.Value ?? related.StyleId?.Value}'.", false, location) { IsDegradation = true }); return metadata with { Styles = new Dictionary<string, string>(metadata.Styles, StringComparer.Ordinal) { ["codeInline"] = related.StyleId?.Value ?? selector } }; }
            diagnostics.Add(new("YADG-YOLO-CODE-001", $"Configured inline-code style '{selector}' is unavailable; fallback: ordinary text.", false, location) { IsDegradation = true });
            return metadata with { Styles = new Dictionary<string, string>(metadata.Styles, StringComparer.Ordinal) { ["codeInline"] = "builtin:plain-v1" } };
        }
        if (selected is null) { diagnostics.Add(new("YADG-WORD-STYLE", $"Style selector '{selector}' does not resolve. Run yadg inspect styles.", true, location)); return metadata; }
        if (selected.Type?.Value != StyleValues.Character) { diagnostics.Add(new("YADG-WORD-STYLE", $"Style selector '{selector}' resolves to {StyleTypeName(selected.Type?.Value)}; role 'codeInline' requires character. Run yadg inspect styles.", true, location)); return metadata; }
        var normalized = new Dictionary<string, string>(metadata.Styles, StringComparer.Ordinal) { ["codeInline"] = selected.StyleId?.Value ?? selector };
        return metadata with { Styles = normalized };
    }

    private static void ValidateList(YadgList list, Styles? styles, Numbering? numbering, TemplateMetadata metadata, string location, List<Diagnostic> diagnostics, bool yolo = false)
    {
        if (StyleFor(metadata, list.Ordered ? "ordered" : "unordered", "").StartsWith("builtin:", StringComparison.Ordinal)) return;
            if (metadata.PrototypeBindings.TryGetValue(list.Ordered ? "orderedListItem" : "unorderedListItem", out var prototypeId))
        {
            if (!metadata.ListItemPrototypes.TryGetValue(prototypeId, out var prototype)) { diagnostics.Add(new("YADG-WORD-LIST", $"Configured list prototype '{prototypeId}' is missing.", true, location)); return; }
            if (ResolveParagraphNumberingId(prototype.Paragraph, styles, numbering) is not int prototypeNum || numbering?.Elements<NumberingInstance>().All(n => n.NumberID?.Value != prototypeNum) != false) diagnostics.Add(new("YADG-WORD-LIST", $"List prototype '{prototypeId}' has no valid effective list numbering.", true, location));
            return;
        }
        var style = StyleFor(metadata, list.Ordered ? "ordered" : "unordered", list.Ordered ? "ListNumber" : "ListBullet");
        if (!HasStyle(styles, style)) { if (yolo) diagnostics.Add(new("YADG-YOLO-LIST-001", $"List style '{style}' is unavailable; fallback: builtin:{(list.Ordered ? "ordered" : "unordered")}-list-v1.", false, location) { IsDegradation = true }); else diagnostics.Add(new("YADG-WORD-LIST", $"Missing required list style '{style}'. Run yadg inspect styles to discover available styles.", true, location)); return; }
        var numId = ResolveNumberingId(styles!, style, numbering);
        if (numId is null || numbering is null || !numbering.Descendants<NumberingInstance>().Any(n => n.NumberID?.Value == numId)) { if (yolo) diagnostics.Add(new("YADG-YOLO-LIST-002", $"List style '{style}' has no usable numbering; fallback: builtin:{(list.Ordered ? "ordered" : "unordered")}-list-v1.", false, location) { IsDegradation = true }); else diagnostics.Add(new("YADG-WORD-LIST", $"List style '{style}' does not resolve to a valid numbering definition. Run yadg inspect styles or bind an unorderedListItem/orderedListItem prototype.", true, location)); }
    }

    private static int? ResolveParagraphNumberingId(Paragraph paragraph, Styles? styles, Numbering? numbering = null)
    {
        var direct = paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        if (direct is not null) return direct;
        var id = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return id is null || styles is null ? null : ResolveNumberingId(styles, id, numbering);
    }

    private static int? ResolveNumberingId(Styles styles, string styleId)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (seen.Add(styleId))
        {
            var style = styles.Descendants<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style is null) return null;
            var direct = style.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            if (direct is not null) return direct;
            styleId = style.BasedOn?.Val?.Value ?? string.Empty;
            if (styleId.Length == 0) return null;
        }
        return null;
    }

    private static int? ResolveNumberingId(Styles styles, string styleId, Numbering? numbering)
    {
        var direct = ResolveNumberingId(styles, styleId);
        if (direct is not null) return direct;
        if (numbering is null) return null;
        var styleChain = new HashSet<string>(StringComparer.Ordinal);
        var chainId = styleId;
        while (chainId.Length > 0 && styleChain.Add(chainId)) chainId = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == chainId)?.BasedOn?.Val?.Value ?? string.Empty;
        var abstractIds = numbering.Elements<NumberingInstance>().Where(instance =>
        {
            var abstractId = instance.AbstractNumId?.Val?.Value;
            var abstractNum = numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId);
            return abstractNum?.Descendants<ParagraphStyleIdInLevel>().Any(p => p.Val?.Value is string linkedStyle && styleChain.Contains(linkedStyle)) == true;
        }).Select(instance => instance.NumberID?.Value).Where(x => x is not null).Distinct().ToArray();
        return abstractIds.Length == 1 ? abstractIds[0] : null;
    }

    private static TemplateMetadata ResolveSelectors(TemplateMetadata metadata, Styles? styles, Numbering? numbering, List<Diagnostic> diagnostics, string location, bool yolo = false)
    {
        if (styles is null && !yolo) return metadata;
        var roles = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in metadata.Styles)
        {
            if (pair.Key == "codeInline") { roles[pair.Key] = pair.Value; continue; }
            if (pair.Key == "unordered" && metadata.PrototypeBindings.ContainsKey("unorderedListItem") || pair.Key == "ordered" && metadata.PrototypeBindings.ContainsKey("orderedListItem")) { roles[pair.Key] = pair.Value; continue; }
            var expected = pair.Key is "generatedTable" ? StyleValues.Table : pair.Key is "codeInline" ? StyleValues.Character : StyleValues.Paragraph;
            var all = styles?.Elements<Style>().ToArray() ?? Array.Empty<Style>();
            var exactId = all.FirstOrDefault(s => string.Equals(s.StyleId?.Value, pair.Value, StringComparison.Ordinal));
            Style? selected = exactId;
            if (selected is null)
            {
                var candidates = all.Where(s => string.Equals(s.StyleName?.Val?.Value, pair.Value, StringComparison.OrdinalIgnoreCase) || (s.Aliases?.Val?.Value ?? "").Split(',').Any(a => string.Equals(a.Trim(), pair.Value, StringComparison.OrdinalIgnoreCase))).ToArray();
                if (candidates.Length > 1 && !yolo) diagnostics.Add(new("YADG-WORD-STYLE", $"Style selector '{pair.Value}' is ambiguous: {string.Join(", ", candidates.Select(s => $"{s.StyleName?.Val?.Value ?? s.StyleId?.Value} (id={s.StyleId?.Value})"))}. Run yadg inspect styles.", true, location));
                else if (candidates.Length == 1) selected = candidates[0];
            }
            if (selected is null || selected.Type?.Value != expected)
            {
                if (yolo)
                {
                    var related = FindRoleCandidate(all, pair.Key, expected, numbering);
                    if (related is not null) { roles[pair.Key] = related.StyleId?.Value ?? pair.Value; diagnostics.Add(new("YADG-YOLO-STYLE-001", $"Configured role '{pair.Key}' selector '{pair.Value}' is unavailable or incompatible; fallback: related template style '{related.StyleName?.Val?.Value ?? related.StyleId?.Value}'.", false, location) { IsDegradation = true }); }
                    else if (pair.Key is "unordered" or "ordered") { roles[pair.Key] = $"builtin:{pair.Key}-list-v1"; diagnostics.Add(new("YADG-YOLO-LIST-001", $"No compatible template list resource resolves role '{pair.Key}'; fallback: builtin:{pair.Key}-list-v1.", false, location) { IsDegradation = true }); }
                    else if (pair.Key.StartsWith("heading", StringComparison.Ordinal)) { roles[pair.Key] = "builtin:plain-v1"; diagnostics.Add(new("YADG-YOLO-STYLE-001", $"Heading role '{pair.Key}' is unavailable; fallback: builtin:plain-v1 with direct outline level.", false, location) { IsDegradation = true }); }
                    else if (pair.Key == "generatedTable") { roles[pair.Key] = "builtin:plain-v1"; diagnostics.Add(new("YADG-YOLO-TABLE-001", "No table style is available; fallback: generated table without a style reference.", false, location) { IsDegradation = true }); }
                    else if (pair.Key == "caption") { roles[pair.Key] = "builtin:plain-v1"; diagnostics.Add(new("YADG-YOLO-CAPTION-001", "Caption style is unavailable; fallback: plain caption paragraph.", false, location) { IsDegradation = true }); }
                    else { roles[pair.Key] = pair.Value; }
                    continue;
                }
                if (selected is not null) diagnostics.Add(new("YADG-WORD-STYLE", $"Style selector '{pair.Value}' resolves to {StyleTypeName(selected.Type?.Value)}; role '{pair.Key}' requires {StyleTypeName(expected)}. Run yadg inspect styles.", true, location));
                else roles[pair.Key] = pair.Value;
                continue;
            }
            var selectedId = selected.StyleId?.Value ?? pair.Value;
            if (yolo && (pair.Key is "unordered" or "ordered") && !ResolveNumberingIdForRole(all, selectedId, numbering, pair.Key == "ordered" ? NumberFormatValues.Decimal : NumberFormatValues.Bullet))
            {
                var related = FindRoleCandidate(all.Where(s => s.StyleId?.Value != selectedId).ToArray(), pair.Key, expected, numbering);
                if (related is not null) { roles[pair.Key] = related.StyleId?.Value ?? pair.Value; diagnostics.Add(new("YADG-YOLO-LIST-003", $"List style '{selectedId}' has no compatible numbering; fallback: related template style '{related.StyleName?.Val?.Value ?? related.StyleId?.Value}'.", false, location) { IsDegradation = true }); }
                else { roles[pair.Key] = $"builtin:{pair.Key}-list-v1"; diagnostics.Add(new("YADG-YOLO-LIST-002", $"List style '{selectedId}' has no compatible numbering; fallback: builtin:{pair.Key}-list-v1.", false, location) { IsDegradation = true }); }
                continue;
            }
            roles[pair.Key] = selectedId;
        }
        return metadata with { Styles = roles };
    }

    private static Style? FindRoleCandidate(Style[] styles, string role, StyleValues expected, Numbering? numbering)
    {
        IEnumerable<Style> eligible = styles.Where(s => s.Type?.Value == expected);
        if (role is "unordered" or "ordered")
        {
            eligible = eligible.Where(s => ResolveNumberingIdForRole(styles, s.StyleId?.Value, numbering, role == "ordered" ? NumberFormatValues.Decimal : NumberFormatValues.Bullet));
        }
        else if (role == "codeInline") eligible = eligible.Where(s => { var hint = (s.StyleName?.Val?.Value ?? "") + " " + (s.Aliases?.Val?.Value ?? "") + " " + (s.StyleId?.Value ?? ""); return hint.Contains("code", StringComparison.OrdinalIgnoreCase) || hint.Contains("source", StringComparison.OrdinalIgnoreCase) || hint.Contains("mono", StringComparison.OrdinalIgnoreCase); });
        else if (role.StartsWith("heading", StringComparison.Ordinal))
        {
            var level = int.Parse(role[7..]) - 1;
            eligible = eligible.Where(s => s.StyleParagraphProperties?.OutlineLevel?.Val?.Value == level || (s.StyleName?.Val?.Value ?? "").Equals("Heading" + (level + 1), StringComparison.OrdinalIgnoreCase));
        }
        else if (role == "generatedTable") { }
        else if (role == "caption") eligible = eligible.Where(s => (s.StyleName?.Val?.Value ?? "").Contains("caption", StringComparison.OrdinalIgnoreCase) || (s.StyleId?.Value ?? "").Contains("caption", StringComparison.OrdinalIgnoreCase));
        else eligible = eligible.Where(s => (s.StyleName?.Val?.Value ?? "").Contains(role, StringComparison.OrdinalIgnoreCase));
        return eligible.OrderBy(s => RoleSpecificHintRank(s, styles, role)).ThenBy(s => NormalizeVisibleName(s.StyleName?.Val?.Value ?? ""), StringComparer.Ordinal).ThenBy(s => s.StyleId?.Value ?? "", StringComparer.Ordinal).FirstOrDefault();
    }

    private static string NormalizeVisibleName(string value) => Regex.Replace(value.Trim(), @"\s+", " ").ToUpperInvariant();

    private static int RoleSpecificHintRank(Style style, Style[] styles, string role)
    {
        if (role is not ("unordered" or "ordered")) return 0;
        var cue = role == "ordered" ? "number" : "bullet";
        var basedOn = style.BasedOn?.Val?.Value;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrEmpty(basedOn) && seen.Add(basedOn))
        {
            var parent = styles.FirstOrDefault(s => s.StyleId?.Value == basedOn);
            if (parent is null) break;
            var parentText = (parent.StyleName?.Val?.Value ?? "") + " " + (parent.StyleId?.Value ?? "") + " " + (parent.Aliases?.Val?.Value ?? "");
            if (parentText.Contains(cue, StringComparison.OrdinalIgnoreCase)) return 0;
            basedOn = parent.BasedOn?.Val?.Value;
        }
        var own = (style.StyleName?.Val?.Value ?? "") + " " + (style.StyleId?.Value ?? "") + " " + (style.Aliases?.Val?.Value ?? "");
        return own.Contains(cue, StringComparison.OrdinalIgnoreCase) ? 1 : 2;
    }

    private static bool ResolveNumberingIdForRole(Style[] styles, string? styleId, Numbering? numbering, NumberFormatValues format)
    {
        if (styleId is null || numbering is null) return false;
        var id = ResolveNumberingId(new Styles(styles.Select(s => (Style)s.CloneNode(true))), styleId, numbering);
        if (id is null) return false;
        var instance = numbering.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == id);
        var abstractNum = numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == instance?.AbstractNumId?.Val?.Value);
        return abstractNum?.Descendants<Level>().Any(l => l.NumberingFormat?.Val?.Value == format) == true;
    }

    private static string StyleTypeName(StyleValues? type) => type == StyleValues.Paragraph ? "paragraph" : type == StyleValues.Character ? "character" : type == StyleValues.Table ? "table" : type == StyleValues.Numbering ? "numbering" : "unknown";

    private static void ValidateTable(YadgTable table, Styles? styles, TemplateMetadata metadata, string location, List<Diagnostic> diagnostics)
    { if (!StyleFor(metadata, "generatedTable", "TableGrid").StartsWith("builtin:", StringComparison.Ordinal) && !HasStyle(styles, StyleFor(metadata, "generatedTable", "TableGrid"))) diagnostics.Add(new("YADG-WORD-TABLE", $"Missing required Word table style '{StyleFor(metadata, "generatedTable", "TableGrid")}'; run yadg inspect styles to discover available styles.", true, location)); if (!string.IsNullOrWhiteSpace(table.Caption) && !metadata.PrototypeBindings.ContainsKey("tableCaption") && !HasStyle(styles, StyleFor(metadata, "caption", "Caption"))) diagnostics.Add(new("YADG-WORD-STYLE", "Missing required Word style for table caption. Run yadg inspect styles to discover available styles.", true, location)); }

    private static void ValidateFigure(YadgFigure figure, Styles? styles, TemplateMetadata metadata, Paragraph anchor, string location, List<Diagnostic> diagnostics, bool yolo = false)
    {
        if (figure.Asset is null) { if (!yolo) diagnostics.Add(new("YADG-FIGURE-008", $"Figure '{figure.Id}' has no validated image asset.", true, location)); return; }
        if (!string.IsNullOrEmpty(figure.AltText) && !metadata.PrototypeBindings.ContainsKey("figureCaption") && !HasStyle(styles, StyleFor(metadata, "caption", "Caption"))) diagnostics.Add(new("YADG-WORD-STYLE", "Missing required Word style for caption.", true, location));
        if (EffectiveWidth(anchor) <= 0) diagnostics.Add(new("YADG-FIGURE-009", $"Cannot determine effective text width for figure '{figure.Id}'.", true, location));
    }

    private static bool HasStyle(Styles? styles, string id) => styles?.Descendants<Style>().Any(s => s.StyleId?.Value == id) == true;
    private static string StyleFor(TemplateMetadata metadata, string role, string fallback) => metadata.Styles.TryGetValue(role, out var value) ? value : fallback;

    private static int EffectiveHeadingLevel(int sourceLevel, int contextLevel, SectionSelection selection, int rootLevel)
        => selection == SectionSelection.Section
            ? contextLevel + 1 + (sourceLevel - rootLevel)
            : contextLevel + (sourceLevel - rootLevel);

    private static IReadOnlyList<YadgBlock> RebaseBlocks(IReadOnlyList<YadgBlock> blocks, TemplatePlacement placement)
    {
        var rootLevel = placement.RootLevel;
        if (rootLevel == 0) return blocks;
        return blocks.Select(block => block is YadgHeading heading
            ? heading with { Level = EffectiveHeadingLevel(heading.Level, placement.ContextLevel, placement.SectionReference!.Selection, rootLevel) }
            : block).ToArray();
    }

    private static void PopulatePreparedTable(PreparedTableBinding binding, YadgTable semantic, IReadOnlyDictionary<string, string> targets, TemplateMetadata metadata)
    {
        var prototype = binding.PrototypeRow;
        foreach (var sourceRow in semantic.Rows)
        {
            var clone = (TableRow)prototype.CloneNode(true);
            var cells = clone.Elements<TableCell>().ToArray();
            for (var i = 0; i < cells.Length && i < sourceRow.Count; i++)
            {
                var paragraph = cells[i].Elements<Paragraph>().Single();
                PopulatePreparedCell(paragraph, sourceRow[i], targets, metadata);
            }
            prototype.InsertBeforeSelf(clone);
        }
        binding.MarkerRow.Remove();
        prototype.Remove();
    }

    private static void PopulatePreparedCell(Paragraph prototype, IReadOnlyList<YadgInline> source, IReadOnlyDictionary<string, string> targets, TemplateMetadata metadata)
    {
        var combined = LogicalText(prototype);
        var start = combined.IndexOf("{{cell}}", StringComparison.Ordinal);
        if (start < 0) return;
        var end = start + "{{cell}}".Length;
        var runs = prototype.Elements<Run>().ToArray();
        Run? baseRun = null;
        var runOffset = 0;
        foreach (var run in runs)
        {
            var runText = string.Concat(run.Descendants<Text>().Select(t => t.Text));
            if (start >= runOffset && start < runOffset + runText.Length) { baseRun = run; break; }
            runOffset += runText.Length;
        }
        var result = new Paragraph();
        if (prototype.ParagraphProperties is not null) result.Append(prototype.ParagraphProperties.CloneNode(true));
        var generated = CreatePreparedRuns(source, targets, baseRun?.RunProperties, metadata);
        var offset = 0;
        var inserted = false;
        foreach (var child in prototype.Elements())
        {
            if (child is not Run run)
            {
                result.Append(child.CloneNode(true));
                continue;
            }
            var text = string.Concat(run.Descendants<Text>().Select(t => t.Text));
            if (text.Length == 0) { result.Append(run.CloneNode(true)); continue; }
            var rangeEnd = offset + text.Length;
            if (rangeEnd <= start || offset >= end)
            {
                result.Append(run.CloneNode(true));
            }
            else
            {
                var before = Math.Max(0, Math.Min(text.Length, start - offset));
                var afterStart = Math.Max(0, Math.Min(text.Length, end - offset));
                if (before > 0) result.Append(TextRun(run, text[..before]));
                if (!inserted) { foreach (var generatedRun in generated) result.Append(generatedRun); inserted = true; }
                if (afterStart < text.Length) result.Append(TextRun(run, text[afterStart..]));
            }
            offset = rangeEnd;
        }
        prototype.RemoveAllChildren();
        prototype.Append(result.Elements().Select(e => e.CloneNode(true)));
    }

    private static Run TextRun(Run source, string value)
    {
        var result = new Run();
        if (source.RunProperties is not null) result.Append(source.RunProperties.CloneNode(true));
        result.Append(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
        return result;
    }

    private static IReadOnlyList<Run> CreatePreparedRuns(IReadOnlyList<YadgInline> source, IReadOnlyDictionary<string, string> targets, RunProperties? baseProperties, TemplateMetadata metadata)
    {
        var temporary = new Paragraph();
        AppendInlines(temporary, source, targets, metadata: metadata);
        var result = new List<Run>();
        foreach (var run in temporary.Elements<Run>())
        {
            var clone = (Run)run.CloneNode(true);
            result.Add(clone);
        }
        if (baseProperties is not null)
        {
            var holder = new Paragraph(result);
            ApplyBaseRunFormatting(holder, baseProperties);
            result = holder.Elements<Run>().Select(r => (Run)r.CloneNode(true)).ToList();
        }
        return result;
    }

    private static void ApplyBaseRunFormatting(Paragraph paragraph, RunProperties? baseProperties)
    {
        if (baseProperties is null) return;
        foreach (var run in paragraph.Elements<Run>())
        {
            var existing = run.RunProperties;
            var combined = (RunProperties)baseProperties.CloneNode(true);
            if (existing?.Bold is not null && combined.Bold is null) combined.Append(existing.Bold.CloneNode(true));
            if (existing?.Italic is not null && combined.Italic is null) combined.Append(existing.Italic.CloneNode(true));
            if (existing?.RunStyle is not null) combined.RunStyle = (RunStyle)existing.RunStyle.CloneNode(true);
            run.RunProperties = combined;
        }
    }

    private static IEnumerable<OpenXmlElement> CreateElements(YadgBlock block, Paragraph anchor, WordprocessingDocument document, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets, bool yolo = false)
    {
        switch (block)
        {
            case YadgHeading heading: yield return CreateParagraph(heading, anchor, metadata, targets); break;
            case YadgParagraph paragraph: yield return CreateParagraph(paragraph, anchor, metadata, targets); break;
            case YadgList list:
                var builtins = new Dictionary<bool, int>();
                foreach (var item in list.Items) yield return CreateListParagraph(item, list.Ordered, metadata, targets, document, yolo, builtins); break;
            case YadgTable table:
                yield return CreateTable(table, metadata, targets);
                if (!string.IsNullOrWhiteSpace(table.Caption)) yield return CreateCaption(table.Id, table.Caption!, "tableCaption", metadata, targets);
                break;
            case YadgFigure figure:
                if (figure.Asset is null && yolo) yield return new Paragraph(new Run(new Text($"[YADG figure \"{figure.Id}\" unavailable: {BoundExcerpt(figure.AssetPath, 90)}{(string.IsNullOrWhiteSpace(figure.AltText) ? "" : " — " + BoundExcerpt(figure.AltText, 90))}]")));
                else yield return CreateFigureParagraph(figure, anchor, document);
                if (!string.IsNullOrEmpty(figure.AltText)) yield return CreateCaption(figure.Id, figure.AltText, "figureCaption", metadata, targets);
                break;
        }
    }

    private static Paragraph CreateParagraph(YadgBlock block, Paragraph anchor, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets)
    {
        var paragraph = new Paragraph();
        if (block is YadgHeading heading)
        {
            var role = StyleFor(metadata, $"heading{heading.Level}", $"Heading{heading.Level}");
            paragraph.Append(role.StartsWith("builtin:", StringComparison.Ordinal) ? new ParagraphProperties(new OutlineLevel { Val = heading.Level - 1 }) : new ParagraphProperties(new ParagraphStyleId { Val = role }));
        }
        else if (anchor.ParagraphProperties is not null) paragraph.Append(anchor.ParagraphProperties.CloneNode(true));
        AppendInlines(paragraph, block is YadgHeading h ? new[] { new YadgText(h.Text) } : ((YadgParagraph)block).Inlines, targets, metadata: metadata);
        if (block is YadgHeading headingWithId && headingWithId.Id is not null && targets.TryGetValue(headingWithId.Id, out var headingBookmark) && headingBookmark.StartsWith("section:", StringComparison.Ordinal)) WrapParagraphBookmark(paragraph, headingBookmark[8..]);
        return paragraph;
    }

    private static Paragraph CreateListParagraph(YadgListItem item, bool ordered, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets, WordprocessingDocument document, bool yolo, Dictionary<bool, int> builtins)
    {
        var key = ordered ? "orderedListItem" : "unorderedListItem";
        if (metadata.PrototypeBindings.TryGetValue(key, out var prototypeId) && metadata.ListItemPrototypes.TryGetValue(prototypeId, out var prototype))
        {
            var clone = (Paragraph)prototype.Paragraph.CloneNode(true);
            ReplaceListItemPlaceholder(clone, item.Inlines, targets, metadata);
            return clone;
        }
        var style = StyleFor(metadata, ordered ? "ordered" : "unordered", ordered ? "ListNumber" : "ListBullet");
        var properties = style.StartsWith("builtin:", StringComparison.Ordinal) && yolo
            ? new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = GetBuiltinNumbering(document, ordered, builtins) }))
            : new ParagraphProperties(new ParagraphStyleId { Val = style });
        var paragraph = new Paragraph(properties); AppendInlines(paragraph, item.Inlines, targets, metadata: metadata); return paragraph;
    }

    private static int GetBuiltinNumbering(WordprocessingDocument document, bool ordered, Dictionary<bool, int> cache)
    {
        if (cache.TryGetValue(ordered, out var existing)) return existing;
        var main = document.MainDocumentPart!;
        var part = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
        var numbering = part.Numbering ??= new Numbering();
        var abstractId = Math.Max(0, numbering.Elements<AbstractNum>().Select(x => (int?)x.AbstractNumberId?.Value).Max() ?? -1) + 1;
        var numId = Math.Max(0, numbering.Elements<NumberingInstance>().Select(x => (int?)x.NumberID?.Value).Max() ?? 0) + 1;
        var level = new Level { LevelIndex = 0 };
        level.Append(new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = ordered ? NumberFormatValues.Decimal : NumberFormatValues.Bullet }, new LevelText { Val = ordered ? "%1." : "•" }, new LevelJustification { Val = LevelJustificationValues.Left });
        numbering.Append(new AbstractNum(level) { AbstractNumberId = abstractId });
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = abstractId }) { NumberID = numId });
        numbering.Save(); cache[ordered] = numId; return numId;
    }

    private static void ReplaceListItemPlaceholder(Paragraph paragraph, IReadOnlyList<YadgInline> inlines, IReadOnlyDictionary<string, string> targets, TemplateMetadata metadata)
    {
        var texts = paragraph.Descendants<Text>().Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph).ToArray();
        var combined = string.Concat(texts.Select(t => t.Text));
        const string marker = "{{item}}";
        var at = combined.IndexOf(marker, StringComparison.Ordinal);
        if (at < 0 || combined.IndexOf(marker, at + marker.Length, StringComparison.Ordinal) >= 0) return;
        var endAt = at + marker.Length; var offset = 0; var startIndex = -1; var endIndex = -1;
        for (var i = 0; i < texts.Length; i++)
        {
            var next = offset + (texts[i].Text?.Length ?? 0);
            if (startIndex < 0 && at >= offset && at < next) startIndex = i;
            if (endAt > offset && endAt <= next) { endIndex = i; break; }
            offset = next;
        }
        if (startIndex < 0 || endIndex < 0) return;
        var startOffset = texts.Take(startIndex).Sum(t => t.Text?.Length ?? 0);
        var endOffset = texts.Take(endIndex).Sum(t => t.Text?.Length ?? 0);
        var startLocal = at - startOffset; var endLocal = endAt - endOffset;
        var startText = texts[startIndex]; var endText = texts[endIndex];
        var startRun = startText.Ancestors<Run>().FirstOrDefault();
        if (startRun is null) return;
        var prefix = (startText.Text ?? string.Empty)[..startLocal];
        var suffix = (endText.Text ?? string.Empty)[endLocal..];
        startText.Text = prefix; startText.Space = SpaceProcessingModeValues.Preserve;
        if (startIndex == endIndex) { }
        else
        {
            for (var i = startIndex + 1; i < endIndex; i++) texts[i].Text = string.Empty;
            endText.Text = suffix; endText.Space = SpaceProcessingModeValues.Preserve;
        }
        var generated = new Paragraph();
        AppendInlines(generated, inlines, targets, startRun.RunProperties?.Italic is not null, startRun.RunProperties?.Bold is not null, metadata);
        ApplyBaseRunFormatting(generated, startRun.RunProperties);
        OpenXmlElement insertionPoint = startRun;
        foreach (var child in generated.Elements<Run>()) { insertionPoint.InsertAfterSelf(child.CloneNode(true)); insertionPoint = insertionPoint.NextSibling()!; }
        if (startIndex == endIndex && suffix.Length > 0)
        {
            var suffixRun = new Run(); if (startRun.RunProperties is not null) suffixRun.Append(startRun.RunProperties.CloneNode(true));
            suffixRun.Append(new Text(suffix) { Space = SpaceProcessingModeValues.Preserve });
            insertionPoint.InsertAfterSelf(suffixRun);
        }
    }

    private static Table CreateTable(YadgTable table, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets)
    {
        var selectedStyle = StyleFor(metadata, "generatedTable", "TableGrid");
        var properties = new TableProperties(new TableLayout { Type = TableLayoutValues.Autofit }, new TableLook { FirstRow = OnOffValue.FromBoolean(true) });
        if (!selectedStyle.StartsWith("builtin:", StringComparison.Ordinal)) properties.Append(new TableStyle { Val = selectedStyle });
        var result = new Table(properties);
        result.Append(new TableGrid(Enumerable.Range(0, table.Header.Count).Select(_ => new GridColumn())));
        var header = new TableRow(new TableRowProperties(new TableHeader())); foreach (var cell in table.Header) header.Append(CreateCell(cell, targets, metadata)); result.Append(header);
        foreach (var row in table.Rows) { var tableRow = new TableRow(); foreach (var cell in row) tableRow.Append(CreateCell(cell, targets, metadata)); result.Append(tableRow); }
        return result;
    }

    private static TableCell CreateCell(IReadOnlyList<YadgInline> inlines, IReadOnlyDictionary<string, string> targets, TemplateMetadata metadata) { var paragraph = new Paragraph(); AppendInlines(paragraph, inlines, targets, metadata: metadata); return new TableCell(paragraph); }

    private static Paragraph CreateFigureParagraph(YadgFigure figure, Paragraph anchor, WordprocessingDocument document)
    {
        var asset = figure.Asset!; var main = document.MainDocumentPart!; var partType = asset.Format == "PNG" ? ImagePartType.Png : ImagePartType.Jpeg; var part = main.AddImagePart(partType);
        using (var stream = File.OpenRead(asset.FullPath)) part.FeedData(stream);
        var relation = main.GetIdOfPart(part); var width = Math.Min((long)asset.WidthPixels * 9525, EffectiveWidth(anchor)); var height = Math.Max(1, (long)asset.HeightPixels * 9525 * width / ((long)asset.WidthPixels * 9525));
        var drawing = new Drawing(new DW.Inline(new DW.Extent { Cx = width, Cy = height }, new DW.DocProperties { Id = 1U, Name = figure.Id }, new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }), new A.Graphic(new A.GraphicData(new PIC.Picture(new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties { Id = 0U, Name = figure.AssetPath }, new PIC.NonVisualPictureDrawingProperties()), new PIC.BlipFill(new A.Blip { Embed = relation }, new A.Stretch(new A.FillRectangle())), new PIC.ShapeProperties(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = width, Cy = height }), new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));
        return new Paragraph(new Run(drawing));
    }

    private static Paragraph CreateCaption(string targetId, string text, string binding, TemplateMetadata metadata, IReadOnlyDictionary<string, string> targets)
    {
        if (metadata.PrototypeBindings.TryGetValue(binding, out var id) && metadata.Prototypes.TryGetValue(id, out var prototype))
        {
            var clone = (Paragraph)prototype.Paragraph.CloneNode(true);
            ReplaceLogicalText(clone, "{{caption}}", text);
            if (targets.TryGetValue(targetId, out var bookmark)) WrapSequenceInBookmark(clone, bookmark);
            return clone;
        }
        var captionStyle = StyleFor(metadata, "caption", "Caption");
        var plain = new Paragraph(captionStyle.StartsWith("builtin:", StringComparison.Ordinal) ? new ParagraphProperties() : new ParagraphProperties(new ParagraphStyleId { Val = captionStyle }));
        AppendInlines(plain, new[] { new YadgText(text) }, targets);
        return plain;
    }

    private static long EffectiveWidth(Paragraph anchor)
    {
        var body = anchor.Ancestors<Body>().FirstOrDefault(); var section = body?.Elements<SectionProperties>().LastOrDefault(); var page = section?.GetFirstChild<PageSize>(); var margins = section?.GetFirstChild<PageMargin>();
        var width = page?.Width?.Value ?? 0; var left = margins?.Left?.Value ?? 0; var right = margins?.Right?.Value ?? 0; return Math.Max(0, ((long)width - left - right) * 635L);
    }

    private static void AppendInlines(Paragraph paragraph, IReadOnlyList<YadgInline> inlines, IReadOnlyDictionary<string, string> targets, bool italic = false, bool bold = false, TemplateMetadata? metadata = null)
    {
        foreach (var inline in inlines) switch (inline)
        {
            case YadgText text: var run = new Run(); if (italic || bold) run.Append(new RunProperties { Italic = italic ? new Italic() : null, Bold = bold ? new Bold() : null }); run.Append(new Text(text.Value) { Space = SpaceProcessingModeValues.Preserve }); paragraph.Append(run); break;
            case YadgLiteralText literal: var literalRun = new Run(); if (italic || bold) literalRun.Append(new RunProperties { Italic = italic ? new Italic() : null, Bold = bold ? new Bold() : null }); literalRun.Append(new Text(literal.Value) { Space = SpaceProcessingModeValues.Preserve }); paragraph.Append(literalRun); break;
            case YadgReference reference:
                if (targets.TryGetValue(reference.Id, out var bookmark)) paragraph.Append(CreateRefRun(bookmark));
                else paragraph.Append(new Run(new Text($"[@{reference.Id}]") { Space = SpaceProcessingModeValues.Preserve }));
                break;
            case YadgHardBreak: paragraph.Append(new Run(new Break())); break;
            case YadgSoftBreak: paragraph.Append(new Run(new Text(" "))); break;
            case YadgCode code:
                var codeRun = new Run(); var codeProps = new RunProperties { RunStyle = metadata?.Styles.TryGetValue("codeInline", out var codeStyle) == true && !codeStyle.StartsWith("builtin:", StringComparison.Ordinal) ? new RunStyle { Val = codeStyle } : null, Italic = italic ? new Italic() : null, Bold = bold ? new Bold() : null };
                if (codeProps.ChildElements.Count > 0) codeRun.Append(codeProps); codeRun.Append(new Text(code.Value) { Space = SpaceProcessingModeValues.Preserve }); paragraph.Append(codeRun); break;
            case YadgEmphasis emphasis: AppendInlines(paragraph, emphasis.Inlines, targets, true, bold, metadata); break;
            case YadgStrong strong: AppendInlines(paragraph, strong.Inlines, targets, italic, true, metadata); break;
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
        var nonEmpty = paragraphs.FirstOrDefault(p => !string.IsNullOrWhiteSpace(ParagraphText(p)));
        if (nonEmpty is null || !string.Equals(ParagraphText(nonEmpty).Trim(), "{{yadg:frontmatter}}", StringComparison.Ordinal)) return TemplateMetadata.Defaults;
        var controls = new List<Paragraph> { nonEmpty };
        var lines = new List<string>(); var end = -1;
        var start = Array.IndexOf(paragraphs, nonEmpty) + 1;
        for (var i = start; i < paragraphs.Length; i++)
        {
            var text = ParagraphText(paragraphs[i]); controls.Add(paragraphs[i]);
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
            if (key is "generatedTable" or "caption" or "codeInline") { if (indent > 2) diagnostics.Add(new("YADG-FM-004", $"Unknown front matter key '{key}'.", true, location)); else styles[key] = value; continue; }
            if (state == "headings" && Regex.IsMatch(key, "^[1-9]$")) { styles[$"heading{key}"] = value; continue; }
            if (state == "lists" && key is "unordered" or "ordered") { styles[key] = value; continue; }
            if (state == "prototypes" && key is "figureCaption" or "tableCaption" or "unorderedListItem" or "orderedListItem") { if (!bindings.TryAdd(key, value)) diagnostics.Add(new("YADG-FM-005", $"Duplicate front matter binding '{key}'.", true, location)); continue; }
            diagnostics.Add(new("YADG-FM-006", $"Unknown front matter key '{key}'.", true, location));
        }
        if (!versionSeen) diagnostics.Add(new("YADG-FM-003", "Front matter must declare version: 1.", true, location));
        var prototypeMap = new Dictionary<string, CaptionPrototype>(StringComparer.Ordinal); var index = end + 1;
        while (index < paragraphs.Length && LogicalText(paragraphs[index]).Trim().StartsWith("{{yadg:prototype:", StringComparison.Ordinal))
        {
            var marker = ParagraphText(paragraphs[index]).Trim(); var id = Regex.Match(marker, @"^\{\{yadg:prototype:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$"); controls.Add(paragraphs[index]); index++;
            if (!id.Success) { diagnostics.Add(new("YADG-PROT-001", $"Malformed prototype marker '{marker}'.", true, location)); continue; }
            var inside = new List<Paragraph>();
            while (index < paragraphs.Length && !ParagraphText(paragraphs[index]).Trim().Equals($"{{{{/yadg:prototype:{id.Groups["id"].Value}}}}}", StringComparison.Ordinal)) { inside.Add(paragraphs[index]); controls.Add(paragraphs[index]); index++; }
            if (index >= paragraphs.Length) { diagnostics.Add(new("YADG-PROT-002", $"Prototype '{id.Groups["id"].Value}' has no closing marker.", true, location)); break; }
            controls.Add(paragraphs[index]); index++;
            if (inside.Count != 1) { diagnostics.Add(new("YADG-PROT-003", $"Prototype '{id.Groups["id"].Value}' must contain exactly one paragraph.", true, location)); continue; }
            var clone = (Paragraph)inside[0].CloneNode(true); var text = LogicalText(clone); var seq = clone.Descendants<SimpleField>().Count() + clone.Descendants<FieldCode>().Count(t => t.Text?.Contains("SEQ", StringComparison.OrdinalIgnoreCase) == true);
            var proto = new CaptionPrototype(id.Groups["id"].Value, clone, seq, Regex.Matches(text, Regex.Escape("{{caption}}")).Count);
            if (!prototypeMap.TryAdd(proto.Id, proto)) diagnostics.Add(new("YADG-PROT-004", $"Duplicate prototype ID '{proto.Id}'.", true, location));
            var listPrototype = bindings.Any(b => b.Key is "unorderedListItem" or "orderedListItem" && b.Value == proto.Id);
            var placeholder = listPrototype ? "{{item}}" : "{{caption}}";
            var placeholderCount = Regex.Matches(text, Regex.Escape(placeholder)).Count;
            if (listPrototype ? placeholderCount != 1 || seq != 0 : seq != 1 || proto.CaptionPlaceholders != 1) diagnostics.Add(new("YADG-PROT-005", listPrototype ? $"List prototype '{proto.Id}' must contain exactly one {{item}} placeholder and no SEQ field." : $"Caption prototype '{proto.Id}' must contain exactly one SEQ field and one {{caption}} placeholder.", true, location));
            if (Regex.Matches(text, @"\{\{[^{}]+\}\}").Cast<Match>().Any(m => !string.Equals(m.Value, placeholder, StringComparison.Ordinal))) diagnostics.Add(new("YADG-PROT-007", $"Prototype '{proto.Id}' contains another YADG placeholder.", true, location));
        }
        if (paragraphs.Skip(index).Any(p => LogicalText(p).Trim().Equals("{{yadg:frontmatter}}", StringComparison.Ordinal))) diagnostics.Add(new("YADG-FM-007", "Duplicate front matter is not supported.", true, location));
        foreach (var binding in bindings.Values) if (!prototypeMap.ContainsKey(binding)) diagnostics.Add(new("YADG-PROT-006", $"Configured prototype '{binding}' does not exist.", true, location));
        var listPrototypes = new Dictionary<string, CaptionPrototype>(StringComparer.Ordinal);
        foreach (var kind in new[] { "unorderedListItem", "orderedListItem" }) if (bindings.TryGetValue(kind, out var id) && prototypeMap.TryGetValue(id, out var prototype)) listPrototypes[id] = prototype;
        return new TemplateMetadata(1, styles, bindings, prototypeMap, controls) { ListItemPrototypes = listPrototypes };
    }

    private static string LogicalText(OpenXmlElement element) => string.Concat(element.Descendants<Text>().Select(t => t.Text));
    private static void RemoveControlRegion(TemplateMetadata metadata) { foreach (var paragraph in metadata.ControlParagraphs.ToArray()) paragraph.Remove(); }

    private static void ReplaceValues(WordprocessingDocument document, IReadOnlyDictionary<string, string> values)
    {
        var main = document.MainDocumentPart!;
        var stories = new List<IEnumerable<Paragraph>> { main.Document.Body?.Descendants<Paragraph>() ?? Array.Empty<Paragraph>() };
        stories.AddRange(main.HeaderParts.Select(p => p.Header.Descendants<Paragraph>()));
        stories.AddRange(main.FooterParts.Select(p => p.Footer.Descendants<Paragraph>()));
        if (main.FootnotesPart?.Footnotes is not null) stories.Add(main.FootnotesPart.Footnotes.Descendants<Paragraph>());
        if (main.EndnotesPart?.Endnotes is not null) stories.Add(main.EndnotesPart.Endnotes.Descendants<Paragraph>());
        if (main.WordprocessingCommentsPart?.Comments is not null) stories.Add(main.WordprocessingCommentsPart.Comments.Descendants<Paragraph>());
        foreach (var paragraph in stories.SelectMany(p => p).Distinct())
        {
            var combined = ParagraphText(paragraph);
            var matches = Regex.Matches(combined, @"\{\{value:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}").Cast<Match>().ToArray();
            foreach (var match in matches.Reverse())
                if (values.TryGetValue(match.Groups["id"].Value, out var replacement)) ReplaceLogicalText(paragraph, match.Value, replacement);
        }
    }

    private static void ReplaceLogicalText(Paragraph paragraph, string oldText, string replacement)
    {
        var texts = paragraph.Descendants<Text>().Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph).ToArray(); var combined = string.Concat(texts.Select(t => t.Text)); var at = combined.IndexOf(oldText, StringComparison.Ordinal); if (at < 0) return;
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

    private static void ValidateReferences(WordprocessingDocument document, string location, YadgDocument model, TemplateMetadata metadata, IReadOnlyList<TemplatePlacement> placements, IReadOnlyList<PreparedTableBinding> prepared, List<Diagnostic> diagnostics)
    {
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles; var numbering = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var targets = BuildTargets(model, placements, prepared, metadata, styles, numbering, diagnostics, location);
        var blocks = placements.SelectMany(p => p.Kind == PlacementKind.Section ? MarkdownDocumentParser.Resolve(model, p.SectionReference!, location).Blocks ?? Array.Empty<YadgBlock>() : p.Kind == PlacementKind.Table ? new YadgBlock[] { model.FindTable(p.Id)! } : new YadgBlock[] { model.FindFigure(p.Id)! }).Concat(prepared.Select(p => (YadgBlock)model.FindTable(p.Id)!));
        foreach (var reference in blocks.SelectMany(ReferencesIn))
        {
            if (model.FindObject(reference.Id) is null) { diagnostics.Add(new("YADG-REF-004", $"Unresolved semantic reference '[@{reference.Id}]'.", true, location)); continue; }
            if (!targets.ContainsKey(reference.Id)) diagnostics.Add(new("YADG-REF-005", $"Semantic reference '[@{reference.Id}]' has no uniquely rendered numbered target in this template.", true, location));
        }
    }

    private static IReadOnlyDictionary<string, string> BuildTargets(YadgDocument model, TemplateAnalysis analysis, WordprocessingDocument document, string location)
    {
        var diagnostics = new List<Diagnostic>(); var result = BuildTargets(model, analysis.Placements, analysis.PreparedBindings ?? Array.Empty<PreparedTableBinding>(), analysis.Metadata, document.MainDocumentPart?.StyleDefinitionsPart?.Styles, document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering, diagnostics, location);
        if (diagnostics.Any(d => d.IsError)) throw new WordAuthoringException(diagnostics.First(d => d.IsError));
        return result;
    }

    private static Dictionary<string, string> BuildTargets(YadgDocument model, IReadOnlyList<TemplatePlacement> placements, IReadOnlyList<PreparedTableBinding> prepared, TemplateMetadata metadata, Styles? styles, Numbering? numbering, List<Diagnostic> diagnostics, string location)
    {
        var direct = placements.Where(p => p.Kind is PlacementKind.Table or PlacementKind.Figure).Select(p => (p.Kind, p.Id)).ToHashSet();
        foreach (var binding in prepared) direct.Add((PlacementKind.Table, binding.Id));
        var counts = new Dictionary<string, int>(StringComparer.Ordinal); var sections = new HashSet<string>(StringComparer.Ordinal);
        foreach (var placement in placements)
        {
            var blocks = placement.Kind == PlacementKind.Section
                ? RebaseBlocks(MarkdownDocumentParser.Resolve(model, placement.SectionReference!, location).Blocks ?? Array.Empty<YadgBlock>(), placement)
                : placement.Kind == PlacementKind.Table ? new YadgBlock[] { model.FindTable(placement.Id)! } : new YadgBlock[] { model.FindFigure(placement.Id)! };
            foreach (var block in blocks)
            {
                if (placement.Kind == PlacementKind.Section && block is YadgHeading heading && placement.SectionReference!.Selection == SectionSelection.Section && heading.Id is not null) { counts[heading.Id] = counts.GetValueOrDefault(heading.Id) + 1; sections.Add(heading.Id); }
                if (block is YadgTable table && (!direct.Contains((PlacementKind.Table, table.Id)) || placement.Kind == PlacementKind.Table)) counts[table.Id] = counts.GetValueOrDefault(table.Id) + 1;
                if (block is YadgFigure figure && (!direct.Contains((PlacementKind.Figure, figure.Id)) || placement.Kind == PlacementKind.Figure)) counts[figure.Id] = counts.GetValueOrDefault(figure.Id) + 1;
            }
        }
        foreach (var binding in prepared) counts[binding.Id] = counts.GetValueOrDefault(binding.Id) + 1;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (id, count) in counts)
        {
            var target = model.FindObject(id); var valid = count == 1;
            if (target is YadgFigure figure) valid &= !string.IsNullOrWhiteSpace(figure.AltText) && metadata.PrototypeBindings.ContainsKey("figureCaption") && metadata.Prototypes.ContainsKey(metadata.PrototypeBindings["figureCaption"]);
            if (target is YadgTable table) valid &= !string.IsNullOrWhiteSpace(table.Caption) && metadata.PrototypeBindings.ContainsKey("tableCaption") && metadata.Prototypes.ContainsKey(metadata.PrototypeBindings["tableCaption"]);
            if (target is YadgSection section)
            {
                var placement = placements.FirstOrDefault(p => p.Kind == PlacementKind.Section && p.Id == id && p.SectionReference!.Selection == SectionSelection.Section);
                var effective = placement is null ? section.Heading.Level : EffectiveHeadingLevel(section.Heading.Level, placement.ContextLevel, placement.SectionReference!.Selection, placement.RootLevel);
                valid &= sections.Contains(id) && HasEffectiveNumbering(effective, metadata, styles, numbering);
            }
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
        var styleId = StyleFor(metadata, $"heading{level}", $"Heading{level}"); return styles is not null && ResolveNumberingId(styles, styleId, numbering) is int id && numbering?.Descendants<NumberingInstance>().Any(n => n.NumberID?.Value == id) == true;
    }

    private static string BookmarkName(string id)
    {
        using var sha = SHA256.Create(); var hash = Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(id)))[..10]; return "yadg_" + Regex.Replace(id, "[^A-Za-z0-9_]", "_") + "_" + hash;
    }
}
