using System.Net;
using System.Text;

namespace Yadg.Core;

public sealed record DiagramCell(int Row, int Column, char Glyph, int Connections);
public sealed record DiagramRectangle(int Left, int Top, int Right, int Bottom, int ZOrder, bool Reconstructed = false, IReadOnlyList<DiagramRectangle>? CoveredBy = null);
public sealed record TextDiagramModel(int Width, int Height, IReadOnlyList<string> Rows, IReadOnlyList<DiagramCell> Cells, IReadOnlyList<DiagramRectangle> Rectangles, string Svg, bool HasStrongGeometry, bool HasInvalidOcclusion);
public sealed record TextDiagramParseResult(TextDiagramModel? Model, string? Error)
{
    public bool Success => Model is not null && Error is null;
}

/// <summary>Parses source-authored character grids into passive SVG without Word/OpenXML dependencies.</summary>
public static class TextDiagramParser
{
    private const int CellWidth = 14;
    private const int CellHeight = 22;
    private const int MaxDimension = 500;
    private const int MaxSourceLength = 200_000;

    private static readonly Dictionary<char, int> UnicodeConnections = new()
    {
        ['─'] = 10, ['│'] = 5, ['┌'] = 6, ['┐'] = 12, ['└'] = 3, ['┘'] = 9,
        ['├'] = 7, ['┤'] = 13, ['┬'] = 14, ['┴'] = 11, ['┼'] = 15,
        ['→'] = 10, ['←'] = 10, ['↑'] = 5, ['↓'] = 5, ['►'] = 10, ['◄'] = 10, ['▲'] = 5, ['▼'] = 5
    };

    // Bits are north/east/south/west. A candidate edge is retained only when its
    // neighbor exposes the reciprocal connection, keeping punctuation in text inert.
    private static readonly Dictionary<char, int> AsciiConnections = new()
    {
        ['-'] = 10, ['|'] = 5, ['+'] = 15, ['>'] = 10, ['<'] = 10, ['^'] = 5, ['v'] = 5
    };

    public static TextDiagramParseResult Parse(string source)
    {
        if (string.IsNullOrWhiteSpace(source)) return new(null, "Diagram source must not be empty.");
        if (source.Length > MaxSourceLength) return new(null, $"Diagram source exceeds {MaxSourceLength} characters.");
        var rows = source.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
        while (rows.Length > 0 && rows[^1].Length == 0) rows = rows[..^1];
        var width = rows.Length == 0 ? 0 : rows.Max(row => row.Length);
        if (width == 0 || rows.Length == 0) return new(null, "Diagram source must not be empty.");
        if (width > MaxDimension || rows.Length > MaxDimension) return new(null, $"Diagram grid dimensions must not exceed {MaxDimension} by {MaxDimension} cells.");

        var raw = new int[rows.Length, width];
        var chars = new char[rows.Length, width];
        for (var r = 0; r < rows.Length; r++)
        for (var c = 0; c < width; c++)
        {
            var glyph = c < rows[r].Length ? rows[r][c] : ' ';
            chars[r, c] = glyph;
            if (UnicodeConnections.TryGetValue(glyph, out var unicode)) raw[r, c] = unicode;
            else if (AsciiConnections.TryGetValue(glyph, out var ascii)) raw[r, c] = ascii;
        }

        var cells = new List<DiagramCell>();
        for (var r = 0; r < rows.Length; r++)
        for (var c = 0; c < width; c++)
        {
            var mask = raw[r, c];
            if (mask == 0) continue;
            var connected = 0;
            if ((mask & 1) != 0 && r > 0 && (raw[r - 1, c] & 4) != 0) connected |= 1;
            if ((mask & 2) != 0 && c + 1 < width && (raw[r, c + 1] & 8) != 0) connected |= 2;
            if ((mask & 4) != 0 && r + 1 < rows.Length && (raw[r + 1, c] & 1) != 0) connected |= 4;
            if ((mask & 8) != 0 && c > 0 && (raw[r, c - 1] & 2) != 0) connected |= 8;
            // ASCII plus signs participate only in actual connected geometry.
            if (chars[r, c] == '+' && BitCount(connected) == 0) continue;
            if (connected != 0 || IsStandaloneArrow(chars[r, c])) cells.Add(new(r, c, chars[r, c], connected));
        }

        var completeRectangles = FindRectangles(chars, rows.Length, width);
        var reconstructed = FindOccludedRectangles(chars, rows.Length, width, completeRectangles, out var hasInvalidOcclusion);
        var rectangles = reconstructed.OrderByDescending(rect => rect.Reconstructed).ThenByDescending(rect => (rect.Right - rect.Left) * (rect.Bottom - rect.Top)).ToArray();
        var substantialUnicode = cells.Count(c => c.Glyph is '─' or '│' or '┌' or '┐' or '└' or '┘' or '├' or '┤' or '┬' or '┴' or '┼') >= 8
            && cells.Select(c => c.Row).Distinct().Count() >= 3;
        var strong = !hasInvalidOcclusion && (rectangles.Length > 0 || (cells.Count(c => c.Connections != 0) >= 4 && HasArrowPath(cells)) || HasBranch(cells) || substantialUnicode);
        var svg = ToSvg(rows, width, rectangles, cells);
        return new(new(width, rows.Length, rows, cells, rectangles, svg, strong, hasInvalidOcclusion), null);
    }

    private static IReadOnlyList<DiagramRectangle> FindRectangles(char[,] grid, int height, int width)
    {
        var result = new List<DiagramRectangle>();
        for (var top = 0; top < height - 2; top++)
        for (var left = 0; left < width - 2; left++)
        {
            if (!IsCorner(grid[top, left], '┌', '+')) continue;
            for (var right = left + 2; right < width; right++)
            {
                if (!IsCorner(grid[top, right], '┐', '+') || !HorizontalEdge(grid, top, left, right)) continue;
                for (var bottom = top + 2; bottom < height; bottom++)
                {
                    if (!IsCorner(grid[bottom, left], '└', '+') || !IsCorner(grid[bottom, right], '┘', '+')) continue;
                    if (!HorizontalEdge(grid, bottom, left, right) || !VerticalEdge(grid, left, top, bottom) || !VerticalEdge(grid, right, top, bottom)) continue;
                    var contained = result.Any(rect => left > rect.Left && top > rect.Top && right < rect.Right && bottom < rect.Bottom);
                    result.Add(new(left, top, right, bottom, contained ? 1 : 0));
                }
            }
        }
        // Larger containing rectangles are painted first, so nested/occluding shapes sit above them.
        return result.OrderByDescending(rect => (rect.Right - rect.Left) * (rect.Bottom - rect.Top)).ToArray();
    }

    private static IReadOnlyList<DiagramRectangle> FindOccludedRectangles(char[,] grid, int height, int width, IReadOnlyList<DiagramRectangle> known, out bool invalid)
    {
        invalid = false;
        var result = known.ToList();
        for (var top = 0; top < height - 2; top++)
        for (var left = 0; left < width - 2; left++)
        {
            if (!IsCorner(grid[top, left], '┌', '+')) continue;
            for (var right = left + 2; right < width; right++)
            {
                if (!IsCorner(grid[top, right], '┐', '+')) continue;
                for (var bottom = top + 2; bottom < height; bottom++)
                {
                    if (!IsCorner(grid[bottom, left], '└', '+')) continue;
                    if (result.Any(rect => rect.Left == left && rect.Top == top && rect.Right == right && rect.Bottom == bottom)) continue;
                    var missing = new List<(int X, int Y)>();
                    if (!InspectEdge((x, y) => grid[y, x], left + 1, right - 1, top, true, missing) ||
                        !InspectEdge((x, y) => grid[y, x], left + 1, right - 1, bottom, true, missing) ||
                        !InspectEdge((x, y) => grid[y, x], top + 1, bottom - 1, left, false, missing) ||
                        !InspectEdge((x, y) => grid[y, x], top + 1, bottom - 1, right, false, missing)) continue;
                    if (!IsCorner(grid[bottom, right], '┘', '+')) missing.Add((right, bottom));
                    if (missing.Count == 0) continue;
                    var covers = new HashSet<int>();
                    var everyGapCoveredExactlyOnce = true;
                    foreach (var (x, y) in missing)
                    {
                        var covering = known.Select((rect, index) => (rect, index))
                            .Where(item => x >= item.rect.Left && x <= item.rect.Right && y >= item.rect.Top && y <= item.rect.Bottom)
                            .Select(item => item.index).ToArray();
                        if (covering.Length != 1) { everyGapCoveredExactlyOnce = false; break; }
                        covers.Add(covering[0]);
                    }
                    if (everyGapCoveredExactlyOnce)
                        result.Add(new(left, top, right, bottom, 0, true, covers.Order().Select(index => known[index]).ToArray()));
                    else if (missing.Count >= 2)
                        invalid = true;
                }
            }
        }
        return result.OrderByDescending(rect => rect.Reconstructed).ThenByDescending(rect => (rect.Right - rect.Left) * (rect.Bottom - rect.Top)).ToArray();
    }

    private static bool InspectEdge(Func<int, int, char> get, int start, int end, int fixedCoordinate, bool horizontal, List<(int X, int Y)> missing)
    {
        var visible = new List<int>();
        var edgeMissing = new List<int>();
        for (var position = start; position <= end; position++)
        {
            var glyph = horizontal ? get(position, fixedCoordinate) : get(fixedCoordinate, position);
            var isEdgeGlyph = horizontal
                ? glyph is '─' or '-' or '+' or '┬' or '┴' or '┼'
                : glyph is '│' or '|' or '+' or '├' or '┤' or '┼';
            if (isEdgeGlyph) visible.Add(position);
            else edgeMissing.Add(position);
        }
        // Missing cells must form a single interior interval with visible fragments on both sides.
        if (edgeMissing.Count == 0) return true;
        if (visible.Count < 2) return false;
        if (edgeMissing.Count != edgeMissing[^1] - edgeMissing[0] + 1) return false;
        missing.AddRange(edgeMissing.Select(position => horizontal ? (position, fixedCoordinate) : (fixedCoordinate, position)));
        return true;
    }

    private static bool IsCorner(char value, char unicode, char ascii) => value == unicode || value == ascii;
    private static bool IsStandaloneArrow(char value) => value is '→' or '←' or '↑' or '↓' or '►' or '◄' or '▲' or '▼';
    private static bool HorizontalEdge(char[,] grid, int row, int left, int right)
    {
        for (var c = left + 1; c < right; c++) if (grid[row, c] is not '─' and not '-' and not '┬' and not '┴' and not '┼' and not '+') return false;
        return true;
    }
    private static bool VerticalEdge(char[,] grid, int col, int top, int bottom)
    {
        for (var r = top + 1; r < bottom; r++) if (grid[r, col] is not '│' and not '|' and not '├' and not '┤' and not '┼' and not '+') return false;
        return true;
    }

    private static bool HasArrowPath(IReadOnlyList<DiagramCell> cells) => cells.Any(cell => cell.Glyph is '→' or '←' or '↑' or '↓' or '►' or '◄' or '▲' or '▼' or '>' or '<' or '^' or 'v');
    private static bool HasBranch(IReadOnlyList<DiagramCell> cells) => cells.Any(cell => cell.Connections is 7 or 11 or 13 or 14 or 15);
    private static int BitCount(int value) => System.Numerics.BitOperations.PopCount((uint)value);

    private static string ToSvg(string[] rows, int width, IReadOnlyList<DiagramRectangle> rectangles, IReadOnlyList<DiagramCell> cells)
    {
        var svgWidth = width * CellWidth + CellWidth;
        var svgHeight = rows.Length * CellHeight + 6;
        var output = new StringBuilder();
        output.Append($"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{svgWidth}\" height=\"{svgHeight}\" viewBox=\"0 0 {svgWidth} {svgHeight}\">");
        foreach (var (rect, index) in rectangles.Select((rect, index) => (rect, index)).Where(item => item.rect.Reconstructed))
        {
            output.Append($"<defs><mask id=\"occlusion-{index}\"><rect width=\"100%\" height=\"100%\" fill=\"white\"/>");
            foreach (var cover in rect.CoveredBy ?? Array.Empty<DiagramRectangle>())
            {
                output.Append($"<rect x=\"{cover.Left * CellWidth + CellWidth / 2}\" y=\"{cover.Top * CellHeight + CellHeight / 2}\" width=\"{(cover.Right - cover.Left) * CellWidth}\" height=\"{(cover.Bottom - cover.Top) * CellHeight}\" fill=\"black\"/>");
            }
            output.Append("</mask></defs>");
        }
        var structural = cells.ToDictionary(cell => (cell.Row, cell.Column));
        output.Append("<g fill=\"none\" stroke=\"#202124\" stroke-width=\"1.5\" stroke-linecap=\"square\" stroke-linejoin=\"round\">");
        foreach (var cell in cells)
        {
            var x = (cell.Column + 0.5) * CellWidth;
            var y = (cell.Row + 0.5) * CellHeight;
            if ((cell.Connections & 2) != 0) output.Append($"<line x1=\"{F(x)}\" y1=\"{F(y)}\" x2=\"{F(x + CellWidth / 2.0)}\" y2=\"{F(y)}\"/>");
            if ((cell.Connections & 4) != 0) output.Append($"<line x1=\"{F(x)}\" y1=\"{F(y)}\" x2=\"{F(x)}\" y2=\"{F(y + CellHeight / 2.0)}\"/>");
            if ((cell.Connections & 8) != 0) output.Append($"<line x1=\"{F(x)}\" y1=\"{F(y)}\" x2=\"{F(x - CellWidth / 2.0)}\" y2=\"{F(y)}\"/>");
            if ((cell.Connections & 1) != 0) output.Append($"<line x1=\"{F(x)}\" y1=\"{F(y)}\" x2=\"{F(x)}\" y2=\"{F(y - CellHeight / 2.0)}\"/>");
            if (cell.Glyph is '→' or '►' or '>') output.Append($"<path d=\"M {F(x + 4)} {F(y - 4)} L {F(x + 8)} {F(y)} L {F(x + 4)} {F(y + 4)}\"/>");
            if (cell.Glyph is '←' or '◄' or '<') output.Append($"<path d=\"M {F(x - 4)} {F(y - 4)} L {F(x - 8)} {F(y)} L {F(x - 4)} {F(y + 4)}\"/>");
            if (cell.Glyph is '↑' or '▲' or '^') output.Append($"<path d=\"M {F(x - 4)} {F(y - 4)} L {F(x)} {F(y - 8)} L {F(x + 4)} {F(y - 4)}\"/>");
            if (cell.Glyph is '↓' or '▼' or 'v') output.Append($"<path d=\"M {F(x - 4)} {F(y + 4)} L {F(x)} {F(y + 8)} L {F(x + 4)} {F(y + 4)}\"/>");
        }
        output.Append("</g>");
        foreach (var (rect, index) in rectangles.Select((rect, index) => (rect, index)))
        {
            var mask = rect.Reconstructed ? $" mask=\"url(#occlusion-{index})\"" : string.Empty;
            output.Append($"<rect x=\"{rect.Left * CellWidth + CellWidth / 2}\" y=\"{rect.Top * CellHeight + CellHeight / 2}\" width=\"{(rect.Right - rect.Left) * CellWidth}\" height=\"{(rect.Bottom - rect.Top) * CellHeight}\" fill=\"none\" stroke=\"#202124\" stroke-width=\"1.5\"{mask}/>");
        }
        for (var r = 0; r < rows.Length; r++)
        {
            var c = 0;
            while (c < rows[r].Length)
            {
                if (structural.ContainsKey((r, c))) { c++; continue; }
                var start = c;
                while (c < rows[r].Length && !structural.ContainsKey((r, c))) c++;
                var value = rows[r][start..c];
                if (string.IsNullOrWhiteSpace(value)) continue;
                output.Append($"<text x=\"{start * CellWidth + CellWidth / 2}\" y=\"{r * CellHeight + CellHeight * 0.72:F1}\" font-family=\"Arial, sans-serif\" font-size=\"14\" fill=\"#202124\">{WebUtility.HtmlEncode(value)}</text>");
            }
        }
        output.Append("</svg>");
        return output.ToString();
    }

    private static bool IsStructuralForDrawing(char value) => UnicodeConnections.ContainsKey(value) || AsciiConnections.ContainsKey(value);
    private static string F(double value) => value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
}
