using System.CommandLine;
using Yadg.Core;
using Yadg.Word;

namespace Yadg.Cli;

public static class Program
{
    public static int Main(string[] args)
    {
        var handlerExitCode = 0;
        var root = new RootCommand("YADG — template-first document authoring");
        root.AddCommand(CreateCheckCommand(() => handlerExitCode = 2));
        root.AddCommand(CreateBuildCommand(() => handlerExitCode = 2));
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
            var diagnostics = ValidateTemplates(loaded);
            PrintDiagnostics(diagnostics);
            if (diagnostics.Any(d => d.IsError)) { fail(); return; }
            Console.WriteLine($"check: valid ({loaded.Sources.Count} source(s), {loaded.Templates.Count} template(s), {loaded.Document.References.Count} reference(s))");
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
            var diagnostics = ValidateTemplates(loaded);
            PrintDiagnostics(diagnostics);
            if (diagnostics.Any(d => d.IsError)) { fail(); return; }
            var output = Path.Combine(loaded.Root, "YadgPreWords");
            Directory.CreateDirectory(output);
            try
            {
                foreach (var template in loaded.Templates)
                    WordAuthoring.Author(template, Path.Combine(output, Path.GetFileName(template)), loaded.Document);
                Console.WriteLine($"build: wrote {loaded.Templates.Count} template output(s) to {output}");
            }
            catch (Exception ex) { Console.Error.WriteLine(new Diagnostic("YADG-BUILD-001", $"Unable to author workspace output: {ex.Message}", true, output)); fail(); }
        }, workspace);
        return command;
    }

    private static IReadOnlyList<Diagnostic> ValidateTemplates(YadgWorkspace workspace)
    {
        var diagnostics = workspace.Diagnostics.ToList();
        foreach (var template in workspace.Templates)
        {
            try { diagnostics.AddRange(WordAuthoring.Analyze(template, workspace.Document).Diagnostics); }
            catch (Exception ex) { diagnostics.Add(new("YADG-WORD-OPEN", $"Cannot inspect DOCX template: {ex.Message}", true, template)); }
        }
        return diagnostics;
    }

    private static void PrintDiagnostics(IEnumerable<Diagnostic> diagnostics)
    {
        foreach (var diagnostic in diagnostics.Where(d => d.IsError)) Console.Error.WriteLine(diagnostic);
    }
}
