namespace Yadg.Core;

public static class AssetValidation
{
    public static YadgDocument AttachAndValidate(YadgDocument document, string workspaceRoot, List<Diagnostic> diagnostics)
    {
        var figures = new Dictionary<string, YadgFigure>(StringComparer.Ordinal);
        foreach (var figure in document.Figures.Values)
        {
            if (figure.GeneratedSource is not null) { figures[figure.Id] = figure; continue; }
            var result = ReadAsset(figure, workspaceRoot);
            if (result.Asset is null) { diagnostics.Add(result.Diagnostic!); figures[figure.Id] = figure; }
            else figures[figure.Id] = figure with { Asset = result.Asset };
        }
        YadgBlock Map(YadgBlock block) => block switch
        {
            YadgFigure figure when figures.TryGetValue(figure.Id, out var attached) => attached,
            _ => block
        };
        var blocks = document.Blocks.Select(Map).ToArray();
        var sections = document.References.ToDictionary(pair => pair.Key, pair => pair.Value with { Body = pair.Value.Body.Select(Map).ToArray() }, StringComparer.Ordinal);
        return document with { Blocks = blocks, References = sections, Figures = figures };
    }

    private static (ImageAsset? Asset, Diagnostic? Diagnostic) ReadAsset(YadgFigure figure, string workspaceRoot)
    {
        if (figure.AssetPath.Contains("://", StringComparison.Ordinal) || figure.AssetPath.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            return (null, new("YADG-FIGURE-003", $"Remote and data URI assets are unsupported: '{figure.AssetPath}'.", true, figure.SourcePath));
        var sourceDirectory = Path.GetDirectoryName(figure.SourcePath) ?? workspaceRoot;
        var fullPath = Path.GetFullPath(Path.Combine(sourceDirectory, figure.AssetPath));
        var relative = Path.GetRelativePath(workspaceRoot, fullPath);
        if (relative.StartsWith("..", StringComparison.Ordinal) || Path.IsPathRooted(relative))
            return (null, new("YADG-FIGURE-004", $"Figure asset must remain within the workspace: '{figure.AssetPath}'.", true, figure.SourcePath));
        if (!File.Exists(fullPath) || File.GetAttributes(fullPath).HasFlag(FileAttributes.ReparsePoint))
            return (null, new("YADG-FIGURE-005", $"Figure asset must be an existing regular file: '{figure.AssetPath}'.", true, figure.SourcePath));
        var extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not ".png" and not ".jpg" and not ".jpeg")
            return (null, new("YADG-FIGURE-006", $"Unsupported figure image format '{extension}'. Only PNG and JPEG are supported.", true, figure.SourcePath));
        try
        {
            var bytes = File.ReadAllBytes(fullPath);
            var dimensions = extension == ".png" ? ReadPng(bytes) : ReadJpeg(bytes);
            if (dimensions is null) return (null, new("YADG-FIGURE-007", $"Image is not a decodable {extension} asset: '{figure.AssetPath}'.", true, figure.SourcePath));
            return (new ImageAsset(fullPath, extension == ".png" ? "PNG" : "JPEG", dimensions.Value.Width, dimensions.Value.Height), null);
        }
        catch (Exception ex) { return (null, new("YADG-FIGURE-007", $"Unable to read figure asset: {ex.Message}", true, figure.SourcePath)); }
    }

    private static (int Width, int Height)? ReadPng(byte[] bytes)
    {
        if (bytes.Length < 24 || bytes[0] != 137 || bytes[1] != 80 || bytes[2] != 78 || bytes[3] != 71) return null;
        return (ReadInt32(bytes, 16), ReadInt32(bytes, 20));
    }

    private static (int Width, int Height)? ReadJpeg(byte[] bytes)
    {
        if (bytes.Length < 4 || bytes[0] != 0xff || bytes[1] != 0xd8) return null;
        var i = 2;
        while (i + 9 < bytes.Length)
        {
            if (bytes[i] != 0xff) { i++; continue; }
            var marker = bytes[i + 1]; i += 2;
            if (marker is 0xd8 or 0xd9) continue;
            if (i + 2 > bytes.Length) return null;
            var length = (bytes[i] << 8) | bytes[i + 1];
            if (length < 2 || i + length > bytes.Length) return null;
            if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf))
                return ((bytes[i + 5] << 8) | bytes[i + 6], (bytes[i + 3] << 8) | bytes[i + 4]);
            i += length;
        }
        return null;
    }

    private static int ReadInt32(byte[] bytes, int offset) => (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];
}
