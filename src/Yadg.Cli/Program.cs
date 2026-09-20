using System.CommandLine;
using Yadg.Core;
using Yadg.Word;
using Yadg.Renderer;

namespace Yadg.Cli;

public static class Program
{
    public static int Main(string[] args)
    {
        var handlerExitCode = 0;
        var root = new RootCommand("YADG — template-first document authoring");
        root.AddCommand(CreateCheckCommand(() => handlerExitCode = 2));
        root.AddCommand(CreateBuildCommand(() => handlerExitCode = 2));
        root.AddCommand(CreateRenderCommand(() => handlerExitCode = 2));
        var commandExitCode = root.Invoke(args);
        return handlerExitCode == 0 ? commandExitCode : handlerExitCode;
    }

    private static Command CreateCheckCommand(Action fail)
    {
        var command = new Command("check", "Validate one YADG workspace without producing outputs.");
        var workspace = new Option<DirectoryInfo?>("--workspace", "Workspace root; defaults to the current directory.");
        command.AddOption(workspace);
        command.SetHandler((DirectoryInfo? path) =>
        {
            var loaded = WorkspaceLoader.Load(path?.FullName);
            try
            {
                var diagnostics = ValidateTemplates(loaded);
                PrintDiagnostics(diagnostics);
                if (diagnostics.Any(d => d.IsError)) { fail(); return; }
                Console.WriteLine($"check: valid ({loaded.Sources.Count} source(s), {loaded.Templates.Count} template(s), {loaded.Document.References.Count} reference(s))");
            }
            finally { loaded.CleanupTemporaryProducerFiles(); }
        }, workspace);
        return command;
    }

    private static Command CreateBuildCommand(Action fail)
    {
        var command = new Command("build", "Build all workspace templates without Microsoft Word.");
        var workspace = new Option<DirectoryInfo?>("--workspace", "Workspace root; defaults to the current directory.");
        command.AddOption(workspace);
        command.SetHandler((DirectoryInfo? path) =>
        {
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
        }, workspace);
        return command;
    }

    private static Command CreateRenderCommand(Action fail)
    {
        var command = new Command("render", "Finalize authored DOCX files through LibreOffice and export PDFs.");
        var workspace = new Option<DirectoryInfo?>("--workspace", "Workspace root; defaults to the current directory.");
        var renderer = new Option<string>("--renderer", () => "libreoffice", "Renderer ID; M0005 supports libreoffice.");
        var rendererPath = new Option<FileInfo?>("--renderer-path", "Explicit LibreOffice soffice executable path.");
        command.AddOption(workspace); command.AddOption(renderer); command.AddOption(rendererPath);
        command.SetHandler((DirectoryInfo? path, string rendererId, FileInfo? executable) =>
        {
            if (!string.Equals(rendererId, "libreoffice", StringComparison.OrdinalIgnoreCase)) { Console.Error.WriteLine("YADG-RENDER-020: Unsupported renderer ID. M0005 supports only 'libreoffice'."); fail(); return; }
            var result = new LibreOfficeRenderer().Render(path?.FullName ?? Directory.GetCurrentDirectory(), executable?.FullName);
            foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine(diagnostic);
            if (!result.Success) { fail(); return; }
            Console.WriteLine($"render: finalized PreWords through LibreOffice ({result.RuntimeVersion}) using {result.Executable}; isolated session {result.ProfileIdentity}");
        }, workspace, renderer, rendererPath);
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
