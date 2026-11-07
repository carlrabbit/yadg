namespace Yadg.Core;

public sealed record Diagnostic(string Code, string Message, bool IsError = true, string? Location = null)
{
    public bool IsDegradation { get; init; }
    public DocxLocation? StructuredLocation { get; init; }
    public Diagnostic(string code, string message, bool isError, DocxLocation location) : this(code, message, isError, location.FilePath) => StructuredLocation = location;
    public override string ToString() => Location is null ? $"{(IsError ? "error" : "warning")} {Code}: {Message}" : $"{(IsError ? "error" : "warning")} {Code} {StructuredLocation?.Format() ?? Location}: {Message}";
}

public sealed record DocxLocation(string FilePath, string Story, IReadOnlyList<string> HeadingContext, int? Ordinal, string Excerpt, string? PackagePart = null, string? ParagraphId = null)
{
    public string Format(string? displayPath = null)
    {
        var context = HeadingContext.Count == 0 ? "" : " > " + string.Join(" > ", HeadingContext.Select(heading => $"\"{heading.Replace("\"", "\\\"", StringComparison.Ordinal)}\""));
        var ordinal = Ordinal is int value ? $" > paragraph {value}" : "";
        return $"{displayPath ?? FilePath}\n{Story}{context}{ordinal}\nnear: \"{Excerpt}\"";
    }
}

public abstract record YadgBlock;
public sealed record YadgHeading(int Level, string Text, string? Id = null) : YadgBlock;
public sealed record YadgParagraph(IReadOnlyList<YadgInline> Inlines) : YadgBlock;
public sealed record YadgList(bool Ordered, IReadOnlyList<YadgListItem> Items, int Level = 0) : YadgBlock;
public sealed record YadgListItem(IReadOnlyList<YadgInline> Inlines, IReadOnlyList<YadgList>? NestedLists = null, bool? TaskChecked = null);
public sealed record YadgTable(string Id, IReadOnlyList<IReadOnlyList<YadgInline>> Header, IReadOnlyList<IReadOnlyList<IReadOnlyList<YadgInline>>> Rows, string? Caption = null) : YadgBlock
{
    public IReadOnlyList<YadgTableCell>? GridHeader { get; init; }
    public IReadOnlyList<IReadOnlyList<YadgTableCell>>? GridRows { get; init; }
    public bool IsGridTable { get; init; }
}
public sealed record YadgTableCell(IReadOnlyList<IReadOnlyList<YadgInline>> Paragraphs, int ColumnSpan = 1, int RowSpan = 1, int StartColumn = 0);
public sealed record YadgCodeBlock(string Text, string? Language = null) : YadgBlock;
public sealed record YadgBlockQuote(IReadOnlyList<YadgBlock> Blocks) : YadgBlock;
public sealed record YadgFigure(string Id, string AltText, string AssetPath, string SourcePath, ImageAsset? Asset = null, string? GeneratedSource = null) : YadgBlock
{
    public string? GeneratedSvg { get; init; }
}
public sealed record YadgDiagram(string SvgContent, string SourcePath) : YadgBlock;

public sealed record ImageAsset(string FullPath, string Format, int WidthPixels, int HeightPixels);

public abstract record YadgInline;
public sealed record YadgText(string Value) : YadgInline;
public sealed record YadgLiteralText(string Value) : YadgInline;
public sealed record YadgEmphasis(IReadOnlyList<YadgInline> Inlines) : YadgInline;
public sealed record YadgStrong(IReadOnlyList<YadgInline> Inlines) : YadgInline;
public sealed record YadgHardBreak : YadgInline;
public sealed record YadgSoftBreak : YadgInline;
public sealed record YadgReference(string Id) : YadgInline;
public sealed record YadgLink(string Destination, IReadOnlyList<YadgInline> Inlines, string? InternalTargetId = null) : YadgInline;
public sealed record YadgCode(string Value) : YadgInline;
public sealed record YadgStrike(IReadOnlyList<YadgInline> Inlines) : YadgInline;

public sealed record ReferenceFormatting(string Headings = "number", string Figures = "number", string Tables = "number")
{
    public static ReferenceFormatting Default { get; } = new();
    public string For(object? target) => target switch
    {
        YadgSection => Headings,
        YadgFigure => Figures,
        YadgTable => Tables,
        _ => "number"
    };
}

public sealed record YadgSection(string Id, YadgHeading Heading, IReadOnlyList<YadgBlock> Body)
{
    public IReadOnlyList<YadgBlock> Select(SectionSelection selection) => selection == SectionSelection.Section
        ? new[] { Heading }.Concat(Body).ToArray()
        : Body;
}

public sealed record YadgDocument(
    IReadOnlyList<YadgBlock> Blocks,
    IReadOnlyDictionary<string, YadgSection> References,
    IReadOnlyDictionary<string, YadgTable> Tables,
    IReadOnlyDictionary<string, YadgFigure> Figures)
{
    public ReferenceFormatting ReferenceFormatting { get; init; } = ReferenceFormatting.Default;
    public YadgSection? FindSection(string id) => References.TryGetValue(id, out var section) ? section : null;
    public YadgTable? FindTable(string id) => Tables.TryGetValue(id, out var table) ? table : null;
    public YadgFigure? FindFigure(string id) => Figures.TryGetValue(id, out var figure) ? figure : null;
    public object? FindObject(string id) => FindSection(id) ?? (object?)FindTable(id) ?? FindFigure(id);
}

public enum SectionSelection { Section, Content }

public sealed record SectionReference(string Id, SectionSelection Selection)
{
    public static bool TryParseTag(string value, out SectionReference? reference, out Diagnostic? diagnostic)
    {
        reference = null;
        diagnostic = null;
        var match = System.Text.RegularExpressions.Regex.Match(value, @"^\{\{(content|section):([A-Za-z][A-Za-z0-9_-]*)\}\}$");
        if (!match.Success)
        {
            diagnostic = new("YADG-TAG-001", $"Malformed or unsupported tag '{value}'. Expected {{content:<stable-id>}} or {{section:<stable-id>}}.");
            return false;
        }
        reference = new(match.Groups[2].Value, match.Groups[1].Value == "section" ? SectionSelection.Section : SectionSelection.Content);
        return true;
    }
}

public sealed record ParseResult(YadgDocument? Document, IReadOnlyList<Diagnostic> Diagnostics)
{
    public bool IsValid => Document is not null && Diagnostics.All(d => !d.IsError);
}
