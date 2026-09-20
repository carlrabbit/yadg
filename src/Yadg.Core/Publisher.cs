namespace Yadg.Core;

public sealed record PublishResult(bool Success, IReadOnlyList<Diagnostic> Diagnostics, string? Destination = null, int PublishedCount = 0);

public static class Publisher
{
    public static PublishResult Publish(YadgWorkspace workspace, string? cliPath = null)
    {
        var diagnostics = workspace.Diagnostics.Where(d => d.IsError).ToList();
        var configured = string.IsNullOrWhiteSpace(cliPath) ? workspace.Values.PublishPath : cliPath;
        if (string.IsNullOrWhiteSpace(configured))
            diagnostics.Add(new("YADG-PUBLISH-001", "No effective publication destination was provided. Use --publish-path or YADG.md publish.path.", true, workspace.Root));
        var destination = configured is null ? null : Path.GetFullPath(Path.IsPathRooted(configured) ? configured : Path.Combine(workspace.Root, configured));
        if (destination is not null && IsReserved(destination, workspace.Root))
            diagnostics.Add(new("YADG-PUBLISH-002", $"Publication destination is reserved or underneath a reserved YADG directory: '{destination}'.", true, destination));

        var source = Path.Combine(workspace.Root, "YadgWords");
        if (!Directory.Exists(source)) diagnostics.Add(new("YADG-PUBLISH-003", $"Finalized DOCX directory does not exist: '{source}'.", true, source));
        var inputs = Directory.Exists(source) ? Directory.EnumerateFiles(source, "*.docx", SearchOption.TopDirectoryOnly).OrderBy(p => p, StringComparer.Ordinal).ToArray() : Array.Empty<string>();
        if (inputs.Length == 0) diagnostics.Add(new("YADG-PUBLISH-004", $"No top-level finalized DOCX files were found in '{source}'.", true, source));
        foreach (var input in inputs) if ((File.GetAttributes(input) & FileAttributes.Directory) != 0) diagnostics.Add(new("YADG-PUBLISH-005", $"Publish input is not a regular file: '{input}'.", true, input));
        if (destination is not null && File.Exists(destination)) diagnostics.Add(new("YADG-PUBLISH-006", $"Publication destination is an existing file: '{destination}'.", true, destination));
        if (diagnostics.Any(d => d.IsError)) return new(false, diagnostics, destination);

        try
        {
            Directory.CreateDirectory(destination!);
            var published = 0;
            foreach (var input in inputs)
            {
                var final = Path.Combine(destination!, Path.GetFileName(input));
                var temp = Path.Combine(destination!, ".yadg-publish-" + Guid.NewGuid().ToString("N") + ".tmp");
                try
                {
                    File.Copy(input, temp, true);
                    File.Move(temp, final, true);
                    published++;
                }
                catch (Exception ex)
                {
                    diagnostics.Add(new("YADG-PUBLISH-007", $"Unable to publish '{input}' to '{final}': {ex.Message}", true, input));
                    try { if (File.Exists(temp)) File.Delete(temp); } catch { }
                    return new(false, diagnostics, destination, published);
                }
            }
            return new(true, diagnostics, destination, published);
        }
        catch (Exception ex)
        {
            diagnostics.Add(new("YADG-PUBLISH-008", $"Unable to prepare publication destination '{destination}': {ex.Message}", true, destination));
            return new(false, diagnostics, destination);
        }
    }

    private static bool IsReserved(string candidate, string root)
    {
        var full = Path.GetFullPath(candidate).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        foreach (var name in new[] { "YadgTemplates", "YadgPreWords", "YadgWords", "YadgPdfs" })
        {
            var reserved = Path.Combine(root, name).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (string.Equals(full, reserved, StringComparison.OrdinalIgnoreCase) || full.StartsWith(reserved + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }
}
