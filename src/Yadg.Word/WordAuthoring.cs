using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Yadg.Core;

namespace Yadg.Word;

public sealed class WordAuthoringException(Diagnostic diagnostic) : Exception(diagnostic.ToString())
{
    public Diagnostic Diagnostic { get; } = diagnostic;
}

public sealed record TemplateAnalysis(IReadOnlyList<(Paragraph Anchor, SectionReference Reference)> Replacements, IReadOnlyList<Diagnostic> Diagnostics)
{
    public bool IsValid => Diagnostics.All(d => !d.IsError);
}

public static class WordAuthoring
{
    private static readonly Regex TagLike = new(@"\{\{[^{}]*\}\}", RegexOptions.Compiled);
    private static readonly Regex ValueTag = new(@"^\{\{value:[A-Za-z][A-Za-z0-9_-]*\}\}$", RegexOptions.Compiled);
    private static readonly Regex SupportedTag = new(@"^\{\{(content|section):([A-Za-z][A-Za-z0-9_-]*)\}\}$", RegexOptions.Compiled);

    public static TemplateAnalysis Analyze(string templatePath, YadgDocument model)
    {
        using var document = WordprocessingDocument.Open(templatePath, false);
        return AnalyzeOpen(document, templatePath, model);
    }

    private static TemplateAnalysis AnalyzeOpen(WordprocessingDocument document, string templatePath, YadgDocument model)
    {
        var diagnostics = new List<Diagnostic>();
        var replacements = new List<(Paragraph Anchor, SectionReference Reference)>();
        var main = document.MainDocumentPart;
        if (main?.Document.Body is null)
        {
            diagnostics.Add(new("YADG-WORD-001", "DOCX has no main document body.", true, templatePath));
            return new(replacements, diagnostics);
        }

        var styles = main.StyleDefinitionsPart?.Styles;
        AnalyzeParagraphs(main.Document.Body.Descendants<Paragraph>(), templatePath, model, replacements, diagnostics, true, styles);
        foreach (var header in main.HeaderParts) AnalyzeParagraphs(header.Header.Descendants<Paragraph>(), templatePath, model, replacements, diagnostics, false, styles);
        foreach (var footer in main.FooterParts) AnalyzeParagraphs(footer.Footer.Descendants<Paragraph>(), templatePath, model, replacements, diagnostics, false, styles);
        return new(replacements, diagnostics);
    }

    public static void Author(string templatePath, string outputPath, YadgDocument model, TemplateAnalysis? analysis = null)
    {
        analysis ??= Analyze(templatePath, model);
        if (!analysis.IsValid) throw new WordAuthoringException(analysis.Diagnostics.First(d => d.IsError));
        File.Copy(templatePath, outputPath, true);
        using var document = WordprocessingDocument.Open(outputPath, true);
        var outputAnalysis = AnalyzeOpen(document, outputPath, model);
        if (!outputAnalysis.IsValid) throw new WordAuthoringException(outputAnalysis.Diagnostics.First(d => d.IsError));
        foreach (var (anchor, reference) in outputAnalysis.Replacements.Reverse())
        {
            var resolved = MarkdownDocumentParser.Resolve(model, reference, outputPath);
            if (resolved.Diagnostic is not null) throw new WordAuthoringException(resolved.Diagnostic);
            var blocks = resolved.Blocks!;
            var generated = blocks.Select(block => CreateParagraph(block, anchor)).ToArray();
            foreach (var paragraph in generated) anchor.InsertBeforeSelf(paragraph);
            anchor.Remove();
        }
        document.MainDocumentPart!.Document.Save();
    }

    private static void AnalyzeParagraphs(IEnumerable<Paragraph> paragraphs, string templatePath, YadgDocument model, List<(Paragraph Anchor, SectionReference Reference)> replacements, List<Diagnostic> diagnostics, bool supportedMainBody, Styles? styles)
    {
        foreach (var paragraph in paragraphs)
        {
            var logical = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
            var matches = TagLike.Matches(logical).Cast<Match>().ToArray();
            if (matches.Length == 0) continue;
            if (!supportedMainBody || paragraph.Parent is not Body || paragraph.Ancestors<Table>().Any())
            {
                diagnostics.Add(new("YADG-WORD-LOCATION", "YADG tags are supported only as standalone paragraphs in the main document body.", true, templatePath));
                continue;
            }
            foreach (var match in matches)
            {
                var trimmed = logical.Trim();
                if (!string.Equals(trimmed, match.Value, StringComparison.Ordinal))
                {
                    diagnostics.Add(new("YADG-WORD-BLOCK", $"Block tag '{match.Value}' must be the only non-whitespace content in its paragraph.", true, templatePath));
                    continue;
                }
                if (ValueTag.IsMatch(match.Value))
                {
                    diagnostics.Add(new("YADG-VALUE-UNSUPPORTED", $"Reserved value tag '{match.Value}' is not supported in M0002.", true, templatePath));
                    continue;
                }
                if (!SupportedTag.IsMatch(match.Value))
                {
                    diagnostics.Add(new("YADG-TAG-001", $"Malformed or obsolete tag '{match.Value}'. Use {{content:<stable-id>}} or {{section:<stable-id>}}.", true, templatePath));
                    continue;
                }
                if (!SectionReference.TryParseTag(match.Value, out var reference, out var tagDiagnostic)) { diagnostics.Add(tagDiagnostic! with { Location = templatePath }); continue; }
                var resolved = MarkdownDocumentParser.Resolve(model, reference!, templatePath);
                if (resolved.Diagnostic is not null) diagnostics.Add(resolved.Diagnostic);
                else
                {
                    replacements.Add((paragraph, reference!));
                    foreach (var heading in resolved.Blocks!.OfType<YadgHeading>())
                    {
                        var styleId = $"Heading{heading.Level}";
                        if (styles is null || !styles.Descendants<Style>().Any(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal)))
                            diagnostics.Add(new("YADG-WORD-STYLE", $"Selected heading level {heading.Level} requires existing Word style '{styleId}'.", true, templatePath));
                    }
                }
            }
        }
    }

    private static Paragraph CreateParagraph(YadgBlock block, Paragraph anchor)
    {
        var paragraph = new Paragraph();
        if (block is YadgHeading heading)
            paragraph.Append(new ParagraphProperties(new ParagraphStyleId { Val = $"Heading{heading.Level}" }));
        else if (anchor.ParagraphProperties is not null)
            paragraph.Append(anchor.ParagraphProperties.CloneNode(true));
        if (block is YadgHeading headingBlock) AppendInlines(paragraph, new[] { new YadgText(headingBlock.Text) });
        if (block is YadgParagraph paragraphBlock) AppendInlines(paragraph, paragraphBlock.Inlines);
        return paragraph;
    }

    private static void AppendInlines(Paragraph paragraph, IReadOnlyList<YadgInline> inlines, bool italic = false, bool bold = false)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case YadgText text:
                    var run = new Run();
                    if (italic || bold) run.Append(new RunProperties { Italic = italic ? new Italic() : null, Bold = bold ? new Bold() : null });
                    run.Append(new Text(text.Value) { Space = SpaceProcessingModeValues.Preserve }); paragraph.Append(run); break;
                case YadgHardBreak: paragraph.Append(new Run(new Break())); break;
                case YadgSoftBreak: paragraph.Append(new Run(new Text(" "))); break;
                case YadgEmphasis emphasis: AppendInlines(paragraph, emphasis.Inlines, true || italic, bold); break;
                case YadgStrong strong: AppendInlines(paragraph, strong.Inlines, italic, true || bold); break;
            }
        }
    }
}
