namespace Yadg.Core;

public sealed record YadgWorkspace(string Root, IReadOnlyList<string> Sources, IReadOnlyList<string> Templates, YadgDocument Document, WorkspaceValues Values, IReadOnlyList<Diagnostic> Diagnostics)
{
    public IReadOnlyList<WorkspaceLoader.ResolvedPresentation> Presentations { get; init; } = Array.Empty<WorkspaceLoader.ResolvedPresentation>();
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
    public static YadgWorkspace Load(string? requestedRoot = null, bool yolo = false)
    {
        var diagnostics = new List<Diagnostic>();
        var root = Path.GetFullPath(requestedRoot ?? Directory.GetCurrentDirectory());
        if (!Directory.Exists(root))
            return Invalid(root, diagnostics, new("YADG-WS-001", $"Workspace directory does not exist: '{root}'."));
        var marker = Path.Combine(root, "YADG.md");
        var values = WorkspaceValues.Empty;
        if (!File.Exists(marker)) diagnostics.Add(new("YADG-WS-001", $"Missing workspace marker '{marker}'.", true, marker));
        else values = WorkspaceValuesParser.Parse(marker, diagnostics);
        var templateDir = Path.Combine(root, "YadgTemplates");
        if (!Directory.Exists(templateDir)) diagnostics.Add(new("YADG-WS-004", $"Missing template directory '{templateDir}'.", true, templateDir));

        var allFiles = new List<string>();
        Enumerate(root, allFiles, diagnostics);
        foreach (var nested in allFiles.Where(f => !PathEquals(f, marker) && Path.GetFileName(f).Equals("YADG.md", StringComparison.Ordinal)))
            diagnostics.Add(new("YADG-WS-002", "Nested YADG.md workspace markers are not supported.", true, nested));

        var declaredPresentations = ResolvePresentations(root, templateDir, values.Presentations, diagnostics);
        var presentationSources = declaredPresentations.SelectMany(deck => deck.SourcePaths).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var sources = allFiles.Where(f => IsMarkdownSource(f, root) && !presentationSources.Contains(Path.GetFullPath(f))).OrderBy(f => f, StringComparer.Ordinal).ToArray();
        if (yolo)
            foreach (var source in sources)
            {
                var text = File.ReadAllText(source);
                if (values.ThematicBreakPolicy == "error" && System.Text.RegularExpressions.Regex.IsMatch(text, @"(?m)^\s{0,3}([-*_])(?:\s*\1){2,}\s*$")) diagnostics.Add(new("YADG-YOLO-MD-001", "Markdown thematic break omitted; fallback: no thematic-break output.", false, source) { IsDegradation = true });
                if (values.CodeInlinePolicy == "error" && System.Text.RegularExpressions.Regex.IsMatch(text, @"(?<!`)`[^`\r\n]+`(?!`)")) diagnostics.Add(new("YADG-YOLO-MD-002", "Inline code styling is unsupported by the configured Markdown policy; fallback: preserve code text as plain text.", false, source) { IsDegradation = true });
            }
        if (sources.Length == 0 && declaredPresentations.Count == 0) diagnostics.Add(new("YADG-WS-003", "Workspace contains no Markdown source files.", true, root));
        var templates = sources.Length > 0 && Directory.Exists(templateDir)
            ? SafeTopLevelFiles(templateDir, diagnostics).Where(f => Path.GetExtension(f).Equals(".docx", StringComparison.OrdinalIgnoreCase)).OrderBy(f => f, StringComparer.Ordinal).ToArray()
            : Array.Empty<string>();
        if (sources.Length > 0 && templates.Length == 0) diagnostics.Add(new("YADG-WS-005", "Workspace contains no top-level DOCX templates in YadgTemplates.", true, templateDir));

        var parsed = sources.Select(path => (path, MarkdownDocumentParser.Parse(File.ReadAllText(path), path, yolo && values.ThematicBreakPolicy == "error" ? "ignore" : values.ThematicBreakPolicy, yolo && values.CodeInlinePolicy == "error" ? "ignore" : values.CodeInlinePolicy, values.DiagramsAutoDetect)));
        var document = MarkdownDocumentParser.Merge(parsed, diagnostics) with { ReferenceFormatting = values.References };
        document = AssetValidation.AttachAndValidate(document, root, diagnostics, yolo);
        var workspace = new YadgWorkspace(root, sources, templates, document, values, diagnostics);
        return workspace with { Document = MermaidProducer.RenderFigures(workspace.Document, root, values, diagnostics, workspace.TemporaryProducerDirectories, yolo), Presentations = declaredPresentations };
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
            var relative = Path.GetRelativePath(directory, entry);
            var name = Path.GetFileName(relative);
            if (name.StartsWith(".", StringComparison.Ordinal) || IsGeneratedDiscoveryDirectory(name)) continue;
            FileAttributes attributes;
            try { attributes = File.GetAttributes(entry); }
            catch (Exception ex) { diagnostics.Add(new("YADG-WS-008", $"Cannot inspect workspace path: {ex.Message}", true, entry)); continue; }
            if ((attributes & FileAttributes.Directory) != 0)
            {
                if ((attributes & FileAttributes.ReparsePoint) == 0) Enumerate(entry, files, diagnostics);
            }
            else files.Add(entry);
        }
    }

    private static IReadOnlyList<string> SafeTopLevelFiles(string directory, List<Diagnostic> diagnostics)
    {
        var files = new List<string>();
        try
        {
            foreach (var file in Directory.EnumerateFiles(directory))
            {
                files.Add(file);
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
        return !parts.Any(p => p.StartsWith(".", StringComparison.Ordinal) || IsGeneratedDiscoveryDirectory(p));
    }

    private static bool IsGeneratedDiscoveryDirectory(string name) => name.Equals("YadgTemplates", StringComparison.OrdinalIgnoreCase) || name.Equals("YadgPreWords", StringComparison.OrdinalIgnoreCase) || name.Equals("YadgWords", StringComparison.OrdinalIgnoreCase) || name.Equals("YadgPdfs", StringComparison.OrdinalIgnoreCase) || name.Equals("YadgPresentations", StringComparison.OrdinalIgnoreCase);

    private static IReadOnlyList<ResolvedPresentation> ResolvePresentations(string root, string templateDirectory, IReadOnlyList<PresentationDeclaration> declarations, List<Diagnostic> diagnostics)
    {
        var resolved = new List<ResolvedPresentation>();
        var outputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in declarations)
        {
            var template = ResolveContained(root, declaration.Template);
            if (template is null || !IsWithin(templateDirectory, template) || !Path.GetExtension(template).Equals(".pptx", StringComparison.OrdinalIgnoreCase) || !File.Exists(template))
                diagnostics.Add(new("YADG-PRESENTATION-CONFIG-012", $"Presentation '{declaration.Id}' template must be an existing .pptx file under YadgTemplates: '{declaration.Template}'.", true, Path.Combine(root, "YADG.md")));

            var sourcePaths = new List<string>();
            var seenSources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var source in declaration.Sources)
            {
                var full = ResolveContained(root, source);
                if (full is null || !Path.GetExtension(full).Equals(".md", StringComparison.OrdinalIgnoreCase) || !File.Exists(full) || IsExcludedPresentationSource(root, full))
                {
                    diagnostics.Add(new("YADG-PRESENTATION-CONFIG-013", $"Presentation '{declaration.Id}' source must be an existing workspace-relative .md file outside template/generated/dot directories: '{source}'.", true, Path.Combine(root, "YADG.md")));
                    continue;
                }
                if (!seenSources.Add(full)) diagnostics.Add(new("YADG-PRESENTATION-CONFIG-014", $"Presentation '{declaration.Id}' lists source more than once: '{source}'.", true, Path.Combine(root, "YADG.md")));
                else sourcePaths.Add(full);
            }
            if (!outputs.Add(declaration.Output)) diagnostics.Add(new("YADG-PRESENTATION-CONFIG-008", $"Duplicate presentation output basename '{declaration.Output}' (case-insensitive).", true, Path.Combine(root, "YADG.md")));
            if (template is not null && sourcePaths.Count > 0) resolved.Add(new(declaration, template, sourcePaths));
        }
        return resolved;
    }

    private static string? ResolveContained(string root, string path)
    {
        try
        {
            if (Path.IsPathRooted(path) || Uri.TryCreate(path, UriKind.Absolute, out var uri) && !uri.IsFile || path.Contains('*') || path.Contains('?')) return null;
            var full = Path.GetFullPath(Path.Combine(root, path));
            var relative = Path.GetRelativePath(root, full);
            return relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal) || Path.IsPathRooted(relative) ? null : full;
        }
        catch { return null; }
    }

    private static bool IsExcludedPresentationSource(string root, string fullPath)
    {
        var parts = Path.GetRelativePath(root, fullPath).Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return parts.Any(part => part.StartsWith(".", StringComparison.Ordinal) || IsGeneratedDiscoveryDirectory(part));
    }

    private static bool IsWithin(string parent, string child)
    {
        var relative = Path.GetRelativePath(Path.GetFullPath(parent), Path.GetFullPath(child));
        return relative != ".." && !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal) && !Path.IsPathRooted(relative);
    }

    public sealed record ResolvedPresentation(PresentationDeclaration Declaration, string TemplatePath, IReadOnlyList<string> SourcePaths);

    private static bool PathEquals(string left, string right) => string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
}
