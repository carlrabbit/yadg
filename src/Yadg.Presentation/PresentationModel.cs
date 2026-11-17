namespace Yadg.Presentation;

public enum PresentationRegionType { Text, Image, Table }
public sealed record PresentationLayout(string Id, IReadOnlyDictionary<string, PresentationRegionType> Regions);
public abstract record PresentationInline;
public sealed record PresentationText(string Value) : PresentationInline;
public sealed record PresentationEmphasis(IReadOnlyList<PresentationInline> Children, bool Strong, bool Italic = false) : PresentationInline;
public sealed record PresentationStrike(IReadOnlyList<PresentationInline> Children) : PresentationInline;
public sealed record PresentationCode(string Value) : PresentationInline;
public sealed record PresentationLink(string Destination, IReadOnlyList<PresentationInline> Children) : PresentationInline;
public sealed record PresentationBreak(bool Hard) : PresentationInline;

public abstract record PresentationBlock;
public sealed record PresentationParagraph(IReadOnlyList<PresentationInline> Inlines, int? HeadingLevel = null) : PresentationBlock;
public sealed record PresentationList(bool Ordered, IReadOnlyList<PresentationListItem> Items) : PresentationBlock;
public sealed record PresentationListItem(IReadOnlyList<PresentationBlock> Blocks);
public sealed record PresentationImage(string Alt, string Path, string SourcePath, string? GeneratedSource = null, string? Id = null, string? Caption = null, string? FailureReason = null) : PresentationBlock;
public sealed record PresentationTable(IReadOnlyList<IReadOnlyList<PresentationInline>> Header, IReadOnlyList<IReadOnlyList<IReadOnlyList<PresentationInline>>> Rows) : PresentationBlock;
public sealed record PresentationRegion(string Id, PresentationRegionType Type, IReadOnlyList<PresentationBlock> Blocks);
public sealed record PresentationSlide(string Id, string Title, string Layout, string SourcePath, int SourceOrdinal, IReadOnlyList<PresentationRegion> Regions)
{
    public IReadOnlyList<PresentationInline> TitleInlines { get; init; } = Array.Empty<PresentationInline>();
}
public sealed record PresentationDocument(IReadOnlyList<PresentationSlide> Slides, IReadOnlyList<Yadg.Core.Diagnostic> Diagnostics)
{
    public bool IsValid => Diagnostics.All(diagnostic => !diagnostic.IsError);
}

public sealed record PresentationSource(string Path, string Content);

public static class PresentationLayouts
{
    private static PresentationLayout Layout(string id, params (string Id, PresentationRegionType Type)[] regions) =>
        new(id, regions.ToDictionary(pair => pair.Id, pair => pair.Type, StringComparer.Ordinal));

    public static IReadOnlyDictionary<string, PresentationLayout> All { get; } = new[]
    {
        Layout("title", ("content", PresentationRegionType.Text)),
        Layout("section"),
        Layout("content", ("content", PresentationRegionType.Text)),
        Layout("two-columns", ("left", PresentationRegionType.Text), ("right", PresentationRegionType.Text)),
        Layout("three-columns", ("left", PresentationRegionType.Text), ("center", PresentationRegionType.Text), ("right", PresentationRegionType.Text)),
        Layout("two-rows", ("top", PresentationRegionType.Text), ("bottom", PresentationRegionType.Text)),
        Layout("three-rows", ("top", PresentationRegionType.Text), ("middle", PresentationRegionType.Text), ("bottom", PresentationRegionType.Text)),
        Layout("grid-2x2", ("top-left", PresentationRegionType.Text), ("top-right", PresentationRegionType.Text), ("bottom-left", PresentationRegionType.Text), ("bottom-right", PresentationRegionType.Text)),
        Layout("image", ("image", PresentationRegionType.Image)),
        Layout("text-image", ("text", PresentationRegionType.Text), ("image", PresentationRegionType.Image)),
        Layout("image-text", ("image", PresentationRegionType.Image), ("text", PresentationRegionType.Text)),
        Layout("image-above-text", ("image", PresentationRegionType.Image), ("text", PresentationRegionType.Text)),
        Layout("text-above-image", ("text", PresentationRegionType.Text), ("image", PresentationRegionType.Image)),
        Layout("table", ("table", PresentationRegionType.Table)),
        Layout("text-table", ("text", PresentationRegionType.Text), ("table", PresentationRegionType.Table)),
        Layout("table-text", ("table", PresentationRegionType.Table), ("text", PresentationRegionType.Text))
    }.ToDictionary(layout => layout.Id, StringComparer.Ordinal);
}
