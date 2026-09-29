using System.CommandLine;
using System.Reflection;
using System.Runtime.Versioning;
using DocumentFormat.OpenXml.Packaging;
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
        var command = new Command("check", "Validate one YADG workspace; --yolo enables explicit best-effort recoveries.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        var list = new Option<bool>("--list") { Description = "List discovered sources, templates, and references." };
        var yolo = new Option<bool>("--yolo") { Description = "Explicit best-effort authoring with visible, deterministic degradations." };
        command.Add(workspace); command.Add(list); command.Add(yolo);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var useYolo = parseResult.GetValue(yolo); var loaded = WorkspaceLoader.Load(path?.FullName, useYolo);
            try
            {
                var validation = ValidateTemplates(loaded, useYolo);
                var diagnostics = validation.Diagnostics;
                if (parseResult.GetValue(list)) PrintInventory(loaded);
                PrintDiagnostics(diagnostics, loaded.Root);
                var refs = loaded.Document.References.Count + loaded.Document.Tables.Count + loaded.Document.Figures.Count;
                if (diagnostics.Any(d => d.IsError)) { Console.WriteLine($"check: invalid ({loaded.Sources.Count} source(s), {loaded.Templates.Count} template(s), {refs} reference(s))"); fail(); return; }
                if (useYolo && validation.UsableTemplates == 0) { Console.Error.WriteLine("error YADG-YOLO-OUTPUT-001: No usable template remains for a useful build."); fail(); return; }
                var degradeCount = diagnostics.Count(d => d.IsDegradation);
                Console.WriteLine($"check: valid ({loaded.Sources.Count} source(s), {loaded.Templates.Count} template(s), {refs} reference(s)){(useYolo ? $"; degradations={degradeCount}" : "")}");
            }
            finally { loaded.CleanupTemporaryProducerFiles(); }
        });
        return command;
    }

    private static Command CreateBuildCommand(Action fail)
    {
        var command = new Command("build", "Build authored DOCX templates; --yolo enables explicit best-effort presentation recoveries.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        var yolo = new Option<bool>("--yolo") { Description = "Explicit best-effort authoring with visible, deterministic degradations." };
        command.Add(workspace); command.Add(yolo);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var useYolo = parseResult.GetValue(yolo); var loaded = WorkspaceLoader.Load(path?.FullName, useYolo);
            string? stageRoot = null;
            try
            {
                PrintDiagnostics(loaded.Diagnostics, loaded.Root);
                if (loaded.Diagnostics.Any(d => d.IsError)) { fail(); return; }
                var output = Path.Combine(loaded.Root, "YadgPreWords");
                stageRoot = Path.Combine(loaded.Root, ".yadg-build-stage-" + Guid.NewGuid().ToString("N"));
                var stageOutput = Path.Combine(stageRoot, "YadgPreWords");
                Directory.CreateDirectory(stageOutput);
                var written = 0; var degradationCount = loaded.Diagnostics.Count(d => d.IsDegradation);
                foreach (var template in loaded.Templates)
                {
                    TemplateAnalysis analysis;
                    try { analysis = WordAuthoring.Analyze(template, loaded.Document, loaded.Values.Values, useYolo); }
                    catch (Exception ex)
                    {
                        var recoverable = IsRecoverableTemplateFailure(ex);
                        if (useYolo && recoverable)
                        {
                            Console.Error.WriteLine($"degradation YADG-YOLO-TEMPLATE-001 {Path.GetRelativePath(loaded.Root, template)}: file-local access failure; fallback: skip this template. Reason: {ex.Message}");
                            degradationCount++;
                            continue;
                        }
                        PrintDiagnostics(new[] { new Diagnostic("YADG-WORD-OPEN", $"Template inspection failed; no template skip was applied. Reason: {ex.Message}", true, template) }, loaded.Root);
                        fail(); return;
                    }
                    PrintDiagnostics(analysis.Diagnostics, loaded.Root);
                    if (analysis.Diagnostics.Any(d => d.IsError)) { fail(); return; }
                    degradationCount += analysis.Diagnostics.Count(d => d.IsDegradation);
                    WordAuthoring.Author(template, Path.Combine(stageOutput, Path.GetFileName(template)), loaded.Document, loaded.Values.Values, useYolo); written++;
                }
                if (written == 0) { Console.Error.WriteLine("error YADG-YOLO-OUTPUT-001: No requested template produced a useful authored DOCX."); fail(); return; }
                CommitPreWordOutputs(stageOutput, output);
                Console.WriteLine($"build: wrote {written} template output(s) to {output}{(useYolo ? $"; degradations={degradationCount}" : "")}");
            }
            catch (Exception ex) { Console.Error.WriteLine(new Diagnostic("YADG-BUILD-001", $"Unable to author workspace output: {ex.Message}", true, loaded.Root)); fail(); }
            finally
            {
                loaded.CleanupTemporaryProducerFiles();
                if (stageRoot is not null && Directory.Exists(stageRoot)) try { Directory.Delete(stageRoot, true); } catch { }
            }
        });
        return command;
    }

    private static Command CreateRenderCommand(Action fail)
    {
        var command = new Command("render", "Finalize authored DOCX files; --yolo may use the alternate supported renderer.");
        var workspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        var renderer = new Option<string>("--renderer") { Description = "Renderer ID: libreoffice (default) or word.", DefaultValueFactory = _ => "libreoffice" };
        var rendererPath = new Option<FileInfo?>("--renderer-path") { Description = "Explicit LibreOffice soffice executable path." };
        var yolo = new Option<bool>("--yolo") { Description = "Explicit best-effort rendering; may use the alternate supported renderer." };
        command.Add(workspace); command.Add(renderer); command.Add(rendererPath); command.Add(yolo);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var rendererId = parseResult.GetValue(renderer)!;
            var executable = parseResult.GetValue(rendererPath);
            var useYolo = parseResult.GetValue(yolo);
            RenderResult result;
            if (!string.Equals(rendererId, "word", StringComparison.OrdinalIgnoreCase) && !string.Equals(rendererId, "libreoffice", StringComparison.OrdinalIgnoreCase)) { Console.Error.WriteLine($"YADG-RENDER-020: Unsupported renderer ID '{rendererId}'. Supported renderers are 'libreoffice' and 'word'."); fail(); return; }
            if (string.Equals(rendererId, "word", StringComparison.OrdinalIgnoreCase) && executable is not null) { Console.Error.WriteLine("YADG-RENDER-021: --renderer-path is valid only for LibreOffice."); fail(); return; }
            var root = Path.GetFullPath(path?.FullName ?? Directory.GetCurrentDirectory());
            if (useYolo)
            {
                result = RunStagedRenderer(rendererId, root, executable?.FullName);
                if (!result.Success)
                {
                    var alternate = rendererId.Equals("word", StringComparison.OrdinalIgnoreCase) ? "libreoffice" : "word";
                    var second = RunStagedRenderer(alternate, root, null);
                    if (!second.Success) { foreach (var d in second.Diagnostics) Console.Error.WriteLine(d); fail(); return; }
                    Console.Error.WriteLine($"degradation YADG-YOLO-RENDER-001: requested renderer '{rendererId}' failed; fallback: '{second.ActualRenderer}' ({second.RuntimeVersion ?? "runtime version unavailable"}).");
                    foreach (var d in result.Diagnostics) Console.Error.WriteLine($"  detail {d}");
                    result = second;
                }
                var actual = result.ActualRenderer;
                foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine(diagnostic);
                var fallbackUsed = !actual.Equals(rendererId, StringComparison.OrdinalIgnoreCase);
                Console.WriteLine($"render: finalized PreWords; requested={rendererId}; actual={actual}; runtime={result.RuntimeVersion ?? "runtime detected"}; yolo={(fallbackUsed ? "fallback used" : "no fallback")}; degradations={(fallbackUsed ? 1 : 0)}");
                return;
            }
            if (string.Equals(rendererId, "word", StringComparison.OrdinalIgnoreCase)) result = new WordRendererEngine().Render(root);
            else result = new LibreOfficeRenderer().Render(root, executable?.FullName);
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
            var root = Path.GetFullPath(path?.FullName ?? Directory.GetCurrentDirectory());
            var diagnostics = new List<Diagnostic>();
            string? configured = null;
            if (destination is null)
            {
                var marker = Path.Combine(root, "YADG.md");
                if (File.Exists(marker)) configured = WorkspaceValuesParser.ParsePublishPath(marker, diagnostics);
                else diagnostics.Add(new("YADG-PUBLISH-009", $"Missing YADG.md needed to read publish.path: '{marker}'.", true, marker));
            }
            if (diagnostics.Any(d => d.IsError)) { PrintDiagnostics(diagnostics, root); fail(); return; }
            var result = Publisher.Publish(root, destination?.FullName, configured);
            PrintDiagnostics(result.Diagnostics, root);
            if (!result.Success) { fail(); return; }
            Console.WriteLine($"publish: copied {result.PublishedCount} finalized DOCX file(s) to {result.Destination}");
        });
        return command;
    }

    private sealed record TemplateValidation(IReadOnlyList<Diagnostic> Diagnostics, int UsableTemplates);

    private static TemplateValidation ValidateTemplates(YadgWorkspace workspace, bool yolo = false)
    {
        var diagnostics = workspace.Diagnostics.ToList();
        var usable = 0;
        foreach (var template in workspace.Templates)
        {
            try
            {
                var analysis = WordAuthoring.Analyze(template, workspace.Document, workspace.Values.Values, yolo);
                diagnostics.AddRange(analysis.Diagnostics);
                if (analysis.IsValid) usable++;
            }
            catch (Exception ex)
            {
                var recoverable = IsRecoverableTemplateFailure(ex);
                var degradation = yolo && recoverable;
                diagnostics.Add(new(degradation ? "YADG-YOLO-TEMPLATE-001" : "YADG-WORD-OPEN", degradation ? $"Template cannot be read due to a file-local access failure; fallback: skip this template. Reason: {ex.Message}" : $"Template inspection failed ({(recoverable ? "file-local access failure" : "unsafe/corrupt or unexpected failure")}); no template skip was applied. Reason: {ex.Message}", !degradation, template) { IsDegradation = degradation });
            }
        }
        return new(diagnostics, usable);
    }

    private static bool IsRecoverableTemplateFailure(Exception ex)
    {
        if (ex is OpenXmlPackageException or InvalidDataException) return false;
        return ex is FileNotFoundException or DirectoryNotFoundException or UnauthorizedAccessException or IOException;
    }

    private static void CommitPreWordOutputs(string stagedOutput, string outputDirectory)
    {
        Directory.CreateDirectory(outputDirectory);
        var backupDirectory = Path.Combine(Path.GetDirectoryName(stagedOutput)!, "previous-docx");
        Directory.CreateDirectory(backupDirectory);
        var movedPrevious = new List<(string Original, string Backup)>();
        var committed = new List<string>();
        try
        {
            foreach (var existing in Directory.EnumerateFiles(outputDirectory, "*.docx", SearchOption.TopDirectoryOnly).ToArray())
            {
                var backup = Path.Combine(backupDirectory, Path.GetFileName(existing));
                File.Move(existing, backup);
                movedPrevious.Add((existing, backup));
            }
            foreach (var staged in Directory.EnumerateFiles(stagedOutput, "*.docx", SearchOption.TopDirectoryOnly).ToArray())
            {
                var destination = Path.Combine(outputDirectory, Path.GetFileName(staged));
                File.Move(staged, destination);
                committed.Add(destination);
            }
        }
        catch
        {
            foreach (var file in committed) if (File.Exists(file)) File.Delete(file);
            foreach (var (original, backup) in movedPrevious) if (File.Exists(backup)) File.Move(backup, original, true);
            throw;
        }
    }

    private static void PrintDiagnostics(IEnumerable<Diagnostic> diagnostics)
    {
        PrintDiagnostics(diagnostics, null);
    }

    private static void PrintDiagnostics(IEnumerable<Diagnostic> diagnostics, string? workspaceRoot)
    {
        foreach (var diagnostic in diagnostics)
        {
            var location = diagnostic.StructuredLocation is { } structured && workspaceRoot is not null
                ? structured.Format(Path.GetRelativePath(workspaceRoot, structured.FilePath).Replace('\\', '/'))
                : diagnostic.StructuredLocation?.Format() ?? diagnostic.Location;
            var message = diagnostic.Message;
            if (diagnostic.StructuredLocation is null && workspaceRoot is not null && location is not null)
            {
                var root = Path.GetFullPath(workspaceRoot).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                message = message.Replace(root, "", StringComparison.OrdinalIgnoreCase).Replace(root.TrimEnd(Path.DirectorySeparatorChar), ".", StringComparison.OrdinalIgnoreCase);
                var suffix = "";
                var path = location;
                var match = System.Text.RegularExpressions.Regex.Match(location, @"^(.*?)(:\d+(?::\d+)?)$");
                if (match.Success) { path = match.Groups[1].Value; suffix = match.Groups[2].Value; }
                if (Path.IsPathRooted(path) && Path.GetFullPath(path).StartsWith(root, StringComparison.OrdinalIgnoreCase)) location = Path.GetRelativePath(workspaceRoot, path) + suffix;
            }
            var severity = diagnostic.IsDegradation ? "degradation" : diagnostic.IsError ? "error" : "warning";
            Console.Error.WriteLine(location is null ? $"{severity} {diagnostic.Code}: {message}" : $"{severity} {diagnostic.Code} {location}: {message}");
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
        var templateCommand = new Command("template", "Explain YADG presentation roles, examples, locations, and fallback choices.");
        var templateWorkspace = new Option<DirectoryInfo?>("--workspace") { Description = "Workspace root; defaults to the current directory." };
        var templateFile = new Option<string?>("--template") { Description = "One top-level DOCX template filename." };
        templateCommand.Add(templateWorkspace); templateCommand.Add(templateFile);
        templateCommand.SetAction(result =>
        {
            var root = Path.GetFullPath(result.GetValue(templateWorkspace)?.FullName ?? Directory.GetCurrentDirectory());
            var directory = Path.Combine(root, "YadgTemplates"); var requested = result.GetValue(templateFile);
            if (!Directory.Exists(directory)) { Console.Error.WriteLine("error YADG-INSPECT-001: Missing YadgTemplates directory."); fail(); return; }
            var files = Directory.EnumerateFiles(directory, "*.docx", SearchOption.TopDirectoryOnly).OrderBy(Path.GetFileName, StringComparer.Ordinal).ToArray();
            if (requested is not null) files = files.Where(f => string.Equals(Path.GetFileName(f), requested, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (files.Length == 0) { Console.Error.WriteLine("error YADG-INSPECT-001: No matching top-level DOCX template was found."); fail(); return; }
            foreach (var file in files)
                try { foreach (var line in WordAuthoring.InspectTemplate(file, root)) Console.WriteLine(line); }
                catch (Exception ex) { Console.Error.WriteLine($"error YADG-INSPECT-003 {Path.GetRelativePath(root, file)}: {ex.Message}"); fail(); }
        });
        command.Add(styles); command.Add(templateCommand); return command;
    }

    private static RenderResult RunStagedRenderer(string rendererId, string workspaceRoot, string? rendererPath)
    {
        var stage = Path.Combine(workspaceRoot, ".yadg-yolo-stage-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(Path.Combine(stage, "YadgPreWords"));
            foreach (var input in Directory.EnumerateFiles(Path.Combine(workspaceRoot, "YadgPreWords"), "*.docx", SearchOption.TopDirectoryOnly)) File.Copy(input, Path.Combine(stage, "YadgPreWords", Path.GetFileName(input)), true);
            var result = rendererId.Equals("word", StringComparison.OrdinalIgnoreCase) ? new WordRendererEngine().Render(stage) : new LibreOfficeRenderer().Render(stage, rendererPath);
            if (!result.Success) return result;
            FinalizedOutputCommit.Commit(workspaceRoot, Path.Combine(stage, "YadgWords"), result.ActualRenderer == "libreoffice" ? Path.Combine(stage, "YadgPdfs") : null);
            return result;
        }
        catch (Exception ex) { return new(false, new[] { new RendererDiagnostic("YADG-YOLO-STAGE-001", $"Staged renderer output could not be committed safely: {ex.Message}") }, ActualRenderer: rendererId); }
        finally { try { if (Directory.Exists(stage)) Directory.Delete(stage, true); } catch { } }
    }
}
