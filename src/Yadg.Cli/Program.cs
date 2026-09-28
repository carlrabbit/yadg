using System.CommandLine;
using System.Reflection;
using System.Runtime.Versioning;
using Yadg.Core;
using Yadg.Word;
using Yadg.Renderer;
using Yadg.WordRenderer;
using WordRendererEngine = Yadg.WordRenderer.WordRenderer;

namespace Yadg.Cli;

[SupportedOSPlatform("windows")]
public static class Program
{
    public static int Main(string[] args)
    {
        if (args.Length == 1 && (args[0] == "--version" || args[0] == "-v"))
        {
            Console.WriteLine(ProductVersion());
            return 0;
        }
        var handlerExitCode = 0;
        var root = new RootCommand("YADG — template-first document authoring");
        root.Add(CreateInitCommand(() => handlerExitCode = 2));
        root.Add(CreateCheckCommand(() => handlerExitCode = 2));
        root.Add(CreateInspectCommand(() => handlerExitCode = 2));
        root.Add(CreateBuildCommand(() => handlerExitCode = 2));
        root.Add(CreateRenderCommand(() => handlerExitCode = 2));
        root.Add(CreatePublishCommand(() => handlerExitCode = 2));
        var commandExitCode = root.Parse(args).Invoke();
        return handlerExitCode == 0 ? commandExitCode : handlerExitCode;
    }

    private static string ProductVersion() =>
        typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? typeof(Program).Assembly.GetName().Version?.ToString()
        ?? "unknown";

    private static Command CreateCheckCommand(Action fail)
    {
        var command = new Command("check", "Validate one YADG workspace without producing outputs.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        var list = new Option<bool>("--list") { Description = "List discovered sources, templates, and references." };
        command.Add(workspace);
        command.Add(list);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var loaded = WorkspaceLoader.Load(path?.FullName);
            try
            {
                var diagnostics = ValidateTemplates(loaded);
                if (parseResult.GetValue(list)) PrintInventory(loaded);
                PrintDiagnostics(diagnostics, loaded.Root);
                var refs = loaded.Document.References.Count + loaded.Document.Tables.Count + loaded.Document.Figures.Count;
                if (diagnostics.Any(d => d.IsError)) { Console.WriteLine($"check: invalid ({loaded.Sources.Count} source(s), {loaded.Templates.Count} template(s), {refs} reference(s))"); fail(); return; }
                Console.WriteLine($"check: valid ({loaded.Sources.Count} source(s), {loaded.Templates.Count} template(s), {refs} reference(s))");
            }
            finally { loaded.CleanupTemporaryProducerFiles(); }
        });
        return command;
    }

    private static Command CreateBuildCommand(Action fail)
    {
        var command = new Command("build", "Build all workspace templates without Microsoft Word.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        command.Add(workspace);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var loaded = WorkspaceLoader.Load(path?.FullName);
            try
            {
                var diagnostics = ValidateTemplates(loaded);
                PrintDiagnostics(diagnostics, loaded.Root);
                if (diagnostics.Any(d => d.IsError)) { fail(); return; }
                var output = Path.Combine(loaded.Root, "YadgPreWords");
                Directory.CreateDirectory(output);
                foreach (var template in loaded.Templates)
                    WordAuthoring.Author(template, Path.Combine(output, Path.GetFileName(template)), loaded.Document, loaded.Values.Values);
                Console.WriteLine($"build: wrote {loaded.Templates.Count} template output(s) to {output}");
            }
            catch (Exception ex) { Console.Error.WriteLine(new Diagnostic("YADG-BUILD-001", $"Unable to author workspace output: {ex.Message}", true, loaded.Root)); fail(); }
            finally { loaded.CleanupTemporaryProducerFiles(); }
        });
        return command;
    }

    private static Command CreateRenderCommand(Action fail)
    {
        var command = new Command("render", "Finalize authored DOCX files through the selected renderer.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        var renderer = new Option<string>("--renderer") { Description = "Renderer ID: libreoffice (default) or word.", DefaultValueFactory = _ => "libreoffice" };
        var rendererPath = new Option<FileInfo?>("--renderer-path") { Description = "Explicit LibreOffice soffice executable path." };
        command.Add(workspace); command.Add(renderer); command.Add(rendererPath);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var rendererId = parseResult.GetValue(renderer)!;
            var executable = parseResult.GetValue(rendererPath);
            RenderResult result;
            if (string.Equals(rendererId, "word", StringComparison.OrdinalIgnoreCase))
            {
                if (executable is not null) { Console.Error.WriteLine("YADG-RENDER-021: --renderer-path is valid only for LibreOffice."); fail(); return; }
                result = new WordRendererEngine().Render(path?.FullName ?? Directory.GetCurrentDirectory());
            }
            else if (string.Equals(rendererId, "libreoffice", StringComparison.OrdinalIgnoreCase)) result = new LibreOfficeRenderer().Render(path?.FullName ?? Directory.GetCurrentDirectory(), executable?.FullName);
            else { Console.Error.WriteLine($"YADG-RENDER-020: Unsupported renderer ID '{rendererId}'. Supported renderers are 'libreoffice' and 'word'."); fail(); return; }
            foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine(diagnostic);
            if (!result.Success) { fail(); return; }
            Console.WriteLine($"render: finalized PreWords through {rendererId} ({result.RuntimeVersion ?? "runtime detected"})");
        });
        return command;
    }

    private static Command CreatePublishCommand(Action fail)
    {
        var command = new Command("publish", "Copy finalized DOCX files to an explicit delivery destination.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        var publishPath = new Option<DirectoryInfo?>("--publish-path") { Description = "Delivery directory; overrides YADG.md publish.path." };
        command.Add(workspace); command.Add(publishPath);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var destination = parseResult.GetValue(publishPath);
            var loaded = WorkspaceLoader.Load(path?.FullName);
            try
            {
                var result = Publisher.Publish(loaded, destination?.FullName);
                PrintDiagnostics(result.Diagnostics, loaded.Root);
                if (!result.Success) { fail(); return; }
                Console.WriteLine($"publish: copied {result.PublishedCount} finalized DOCX file(s) to {result.Destination}");
            }
            finally { loaded.CleanupTemporaryProducerFiles(); }
        });
        return command;
    }

    private static IReadOnlyList<Diagnostic> ValidateTemplates(YadgWorkspace workspace)
    {
        var diagnostics = workspace.Diagnostics.ToList();
        foreach (var template in workspace.Templates)
        {
            try { diagnostics.AddRange(WordAuthoring.Analyze(template, workspace.Document, workspace.Values.Values).Diagnostics); }
            catch (Exception ex) { diagnostics.Add(new("YADG-WORD-OPEN", $"Cannot inspect DOCX template: {ex.Message}", true, template)); }
        }
        return diagnostics;
    }

    private static void PrintDiagnostics(IEnumerable<Diagnostic> diagnostics)
    {
        PrintDiagnostics(diagnostics, null);
    }

    private static void PrintDiagnostics(IEnumerable<Diagnostic> diagnostics, string? workspaceRoot)
    {
        foreach (var diagnostic in diagnostics)
        {
            var location = diagnostic.Location;
            var message = diagnostic.Message;
            if (workspaceRoot is not null && location is not null)
            {
                var root = Path.GetFullPath(workspaceRoot).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                message = message.Replace(root, "", StringComparison.OrdinalIgnoreCase).Replace(root.TrimEnd(Path.DirectorySeparatorChar), ".", StringComparison.OrdinalIgnoreCase);
                var suffix = "";
                var path = location;
                var match = System.Text.RegularExpressions.Regex.Match(location, @"^(.*?)(:\d+(?::\d+)?)$");
                if (match.Success) { path = match.Groups[1].Value; suffix = match.Groups[2].Value; }
                if (Path.IsPathRooted(path) && Path.GetFullPath(path).StartsWith(root, StringComparison.OrdinalIgnoreCase)) location = Path.GetRelativePath(workspaceRoot, path) + suffix;
            }
            Console.Error.WriteLine(location is null ? $"{(diagnostic.IsError ? "error" : "warning")} {diagnostic.Code}: {message}" : $"{(diagnostic.IsError ? "error" : "warning")} {diagnostic.Code} {location}: {message}");
        }
    }

    private static void PrintInventory(YadgWorkspace workspace)
    {
        static string Rel(string root, string path) => Path.GetRelativePath(root, path).Replace('\\', '/');
        Console.WriteLine("Sources"); foreach (var item in workspace.Sources.OrderBy(x => Rel(workspace.Root, x), StringComparer.Ordinal)) Console.WriteLine($"  {Rel(workspace.Root, item)}");
        Console.WriteLine("Templates"); foreach (var item in workspace.Templates.OrderBy(x => Rel(workspace.Root, x), StringComparer.Ordinal)) Console.WriteLine($"  {Rel(workspace.Root, item)}");
        Console.WriteLine("References");
        foreach (var id in workspace.Document.References.Keys.OrderBy(x => x, StringComparer.Ordinal)) Console.WriteLine($"  {id}  section");
        foreach (var id in workspace.Document.Tables.Keys.OrderBy(x => x, StringComparer.Ordinal)) Console.WriteLine($"  {id}  table");
        foreach (var id in workspace.Document.Figures.Keys.OrderBy(x => x, StringComparer.Ordinal)) Console.WriteLine($"  {id}  figure");
    }

    private static Command CreateInitCommand(Action fail)
    {
        var command = new Command("init", "Create a starter YADG workspace without overwriting existing files.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." }; command.Add(workspace);
        command.SetAction(result =>
        {
            var root = Path.GetFullPath(result.GetValue(workspace)?.FullName ?? Directory.GetCurrentDirectory());
            var owned = new[] { "YADG.md", "content.md", "YadgTemplates" }.Select(name => Path.Combine(root, name)).Where(path => File.Exists(path) || Directory.Exists(path)).ToArray();
            if (owned.Length > 0) { Console.Error.WriteLine($"error YADG-INIT-001: Bootstrap conflicts: {string.Join(", ", owned.Select(path => Path.GetRelativePath(root, path)))}"); fail(); return; }
            try
            {
                Directory.CreateDirectory(root);
                Directory.CreateDirectory(Path.Combine(root, "YadgTemplates"));
                File.WriteAllText(Path.Combine(root, "YADG.md"), "---\nyadg:\n  version: 1\n---\n");
                File.WriteAllText(Path.Combine(root, "content.md"), "# Introduction {#introduction}\n\nAdd your document content here.\n");
                Console.WriteLine($"Initialized workspace at {root}. Add a prepared DOCX template with {{content:introduction}} in YadgTemplates, then run yadg check.");
            }
            catch (Exception ex) { Console.Error.WriteLine($"error YADG-INIT-002: {ex.Message}"); fail(); }
        });
        return command;
    }

    private static Command CreateInspectCommand(Action fail)
    {
        var command = new Command("inspect", "Inspect template resources.");
        var styles = new Command("styles", "List Word styles, aliases, visibility, and numbering information.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        var template = new Option<string?>("--template") { Description = "One top-level DOCX template filename." };
        styles.Add(workspace); styles.Add(template);
        styles.SetAction(result =>
        {
            var root = Path.GetFullPath(result.GetValue(workspace)?.FullName ?? Directory.GetCurrentDirectory());
            var dir = Path.Combine(root, "YadgTemplates"); var requested = result.GetValue(template);
            if (!Directory.Exists(dir)) { Console.Error.WriteLine("error YADG-INSPECT-001: Missing YadgTemplates directory."); fail(); return; }
            var files = Directory.EnumerateFiles(dir, "*.docx", SearchOption.TopDirectoryOnly).OrderBy(Path.GetFileName, StringComparer.Ordinal).ToArray();
            if (requested is not null) files = files.Where(f => string.Equals(Path.GetFileName(f), requested, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (files.Length == 0) { Console.Error.WriteLine("error YADG-INSPECT-001: No matching top-level DOCX template was found."); fail(); return; }
            foreach (var file in files) try { foreach (var line in WordAuthoring.InspectStyles(file)) Console.WriteLine($"{Path.GetFileName(file)}: {line}"); } catch (Exception ex) { Console.Error.WriteLine($"error YADG-INSPECT-002 {Path.GetRelativePath(root, file)}: {ex.Message}"); fail(); }
        });
        command.Add(styles); return command;
    }
}
