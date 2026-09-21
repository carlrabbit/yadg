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
        root.Add(CreateCheckCommand(() => handlerExitCode = 2));
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
        command.Add(workspace);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var loaded = WorkspaceLoader.Load(path?.FullName);
            try
            {
                var diagnostics = ValidateTemplates(loaded);
                PrintDiagnostics(diagnostics);
                if (diagnostics.Any(d => d.IsError)) { fail(); return; }
                Console.WriteLine($"check: valid ({loaded.Sources.Count} source(s), {loaded.Templates.Count} template(s), {loaded.Document.References.Count} reference(s))");
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
                PrintDiagnostics(diagnostics);
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
                PrintDiagnostics(result.Diagnostics);
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
        foreach (var diagnostic in diagnostics.Where(d => d.IsError)) Console.Error.WriteLine(diagnostic);
    }
}
