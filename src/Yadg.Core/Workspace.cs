namespace Yadg.Core;

public sealed record YadgWorkspace(string Root, IReadOnlyList<string> Sources, IReadOnlyList<string> Templates, YadgDocument Document, WorkspaceValues Values, IReadOnlyList<Diagnostic> Diagnostics)
{
    public List<string> TemporaryProducerDirectories { get; } = new();
    public bool IsValid => Diagnostics.All(d => !d.IsError);
    public void CleanupTemporaryProducerFiles()
    {
        foreach (var directory in TemporaryProducerDirectories.Distinct(StringComparer.OrdinalIgnoreCase))
            try { if (Directory.Exists(directory)) Directory.Delete(directory, true); } catch { }
        TemporaryProducerDirectories.Clear();
    }
}

public static class WorkspaceLoader
{
    public static YadgWorkspace Load(string? requestedRoot = null)
    {
        var diagnostics = new List<Diagnostic>();
        var root = Path.GetFullPath(requestedRoot ?? Directory.GetCurrentDirectory());
        if (!Directory.Exists(root))
            return Invalid(root, diagnostics, new("YADG-WS-001", $"Workspace directory does not exist: '{root}'."));
        if (IsReparse(root))
            return Invalid(root, diagnostics, new("YADG-WS-006", "Workspace root must not be a symbolic link or reparse point."));

        var marker = Path.Combine(root, "YADG.md");
        var values = WorkspaceValues.Empty;
        if (!File.Exists(marker) || IsReparse(marker)) diagnostics.Add(new("YADG-WS-001", $"Missing regular workspace marker '{marker}'.", true, marker));
        else values = WorkspaceValuesParser.Parse(marker, diagnostics);
        var templateDir = Path.Combine(root, "YadgTemplates");
        if (!Directory.Exists(templateDir) || IsReparse(templateDir)) diagnostics.Add(new("YADG-WS-004", $"Missing regular template directory '{templateDir}'.", true, templateDir));
        var outputDir = Path.Combine(root, "YadgPreWords");
        if (!Directory.Exists(outputDir)) diagnostics.Add(new("YADG-WS-007", $"Missing output directory '{outputDir}'. Create YadgPreWords before building.", false, outputDir));

        var allFiles = new List<string>();
        Enumerate(root, allFiles, diagnostics);
        foreach (var nested in allFiles.Where(f => !PathEquals(f, marker) && Path.GetFileName(f).Equals("YADG.md", StringComparison.Ordinal)))
            diagnostics.Add(new("YADG-WS-002", "Nested YADG.md workspace markers are not supported.", true, nested));

        var sources = allFiles.Where(f => IsMarkdownSource(f, root)).OrderBy(f => f, StringComparer.Ordinal).ToArray();
        if (sources.Length == 0) diagnostics.Add(new("YADG-WS-003", "Workspace contains no Markdown source files.", true, root));
        var templates = Directory.Exists(templateDir) && !IsReparse(templateDir)
            ? SafeTopLevelFiles(templateDir, diagnostics).Where(f => Path.GetExtension(f).Equals(".docx", StringComparison.OrdinalIgnoreCase)).OrderBy(f => f, StringComparer.Ordinal).ToArray()
            : Array.Empty<string>();
        if (templates.Length == 0) diagnostics.Add(new("YADG-WS-005", "Workspace contains no top-level DOCX templates in YadgTemplates.", true, templateDir));

        var parsed = sources.Select(path => (path, MarkdownDocumentParser.Parse(File.ReadAllText(path), path)));
        var document = MarkdownDocumentParser.Merge(parsed, diagnostics);
        document = AssetValidation.AttachAndValidate(document, root, diagnostics);
        var workspace = new YadgWorkspace(root, sources, templates, document, values, diagnostics);
        return workspace with { Document = MermaidProducer.RenderFigures(workspace.Document, root, values, diagnostics, workspace.TemporaryProducerDirectories) };
    }

    private static YadgWorkspace Invalid(string root, List<Diagnostic> diagnostics, Diagnostic diagnostic)
    {
        diagnostics.Add(diagnostic);
        return new(root, Array.Empty<string>(), Array.Empty<string>(), new YadgDocument(Array.Empty<YadgBlock>(), new Dictionary<string, YadgSection>(), new Dictionary<string, YadgTable>(), new Dictionary<string, YadgFigure>()), WorkspaceValues.Empty, diagnostics);
    }

    private static void Enumerate(string directory, List<string> files, List<Diagnostic> diagnostics)
    {
        IEnumerable<string> entries;
        try { entries = Directory.EnumerateFileSystemEntries(directory); }
        catch (Exception ex) { diagnostics.Add(new("YADG-WS-008", $"Cannot inspect workspace path: {ex.Message}", true, directory)); return; }
        foreach (var entry in entries)
        {
            if (File.Exists(entry))
            {
                if (IsReparse(entry)) diagnostics.Add(new("YADG-WS-006", "Workspace input files must not be symbolic links or reparse points.", true, entry));
                else files.Add(entry);
            }
            else if (Directory.Exists(entry))
            {
                if (IsReparse(entry)) diagnostics.Add(new("YADG-WS-006", "Workspace discovery does not follow symbolic-link or reparse-point directories.", true, entry));
                else Enumerate(entry, files, diagnostics);
            }
        }
    }

    private static IReadOnlyList<string> SafeTopLevelFiles(string directory, List<Diagnostic> diagnostics)
    {
        var files = new List<string>();
        try
        {
            foreach (var file in Directory.EnumerateFiles(directory))
            {
                if (IsReparse(file)) diagnostics.Add(new("YADG-WS-006", "Template inputs must not be symbolic links or reparse points.", true, file));
                else files.Add(file);
            }
        }
        catch (Exception ex) { diagnostics.Add(new("YADG-WS-008", $"Cannot inspect template directory: {ex.Message}", true, directory)); }
        return files;
    }

    private static bool IsMarkdownSource(string path, string root)
    {
        if (!Path.GetExtension(path).Equals(".md", StringComparison.OrdinalIgnoreCase)) return false;
        if (PathEquals(path, Path.Combine(root, "YADG.md"))) return false;
        var relative = Path.GetRelativePath(root, path);
        var parts = relative.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return !parts.Any(p => p.Equals("YadgTemplates", StringComparison.OrdinalIgnoreCase) || p.Equals("YadgPreWords", StringComparison.OrdinalIgnoreCase) || p.StartsWith(".", StringComparison.Ordinal));
    }

    private static bool PathEquals(string left, string right) => string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
    private static bool IsReparse(string path) => File.GetAttributes(path).HasFlag(FileAttributes.ReparsePoint);
}
