using System.CommandLine;
using System.Reflection;
using System.Runtime.Versioning;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using Yadg.Core;
using Yadg.Word;
using Yadg.Renderer;
using Yadg.WordRenderer;
using Yadg.Presentation;
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
        root.Add(CreatePresentationCommand(() => handlerExitCode = 2));
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
                var diagnostics = validation.Diagnostics.Concat(ValidatePresentations(loaded)).ToArray();
                if (parseResult.GetValue(list)) PrintInventory(loaded);
                PrintDiagnostics(diagnostics, loaded.Root);
                var refs = loaded.Document.References.Count + loaded.Document.Tables.Count + loaded.Document.Figures.Count;
                if (diagnostics.Any(d => d.IsError)) { Console.WriteLine($"check: invalid ({loaded.Sources.Count} source(s), {loaded.Templates.Count} template(s), {loaded.Presentations.Count} presentation(s), {refs} reference(s))"); fail(); return; }
                if (useYolo && validation.UsableTemplates == 0 && loaded.Presentations.Count == 0) { Console.Error.WriteLine("error YADG-YOLO-OUTPUT-001: No usable template remains for a useful build."); fail(); return; }
                var degradeCount = diagnostics.Count(d => d.IsDegradation);
                Console.WriteLine($"check: valid ({loaded.Sources.Count} DOCX source(s), {loaded.Templates.Count} DOCX template(s), {loaded.Presentations.Count} presentation(s), {refs} reference(s)){(useYolo ? $"; degradations={degradeCount}" : "")}");
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
                stageRoot = Path.Combine(loaded.Root, ".yadg-build-stage-" + Guid.NewGuid().ToString("N"));
                var stageOutput = Path.Combine(stageRoot, "YadgPreWords");
                var output = Path.Combine(loaded.Root, "YadgPreWords");
                var written = 0; var degradationCount = loaded.Diagnostics.Count(d => d.IsDegradation);
                if (loaded.Sources.Count > 0)
                {
                    Directory.CreateDirectory(stageOutput);
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
                    if (written == 0) { Console.Error.WriteLine("error YADG-BUILD-001: No requested DOCX template produced an authored output."); fail(); return; }
                    CommitPreWordOutputs(stageOutput, output);
                    Console.WriteLine($"build: wrote {written} DOCX template output(s) to {output}{(useYolo ? $"; degradations={degradationCount}" : "")}");
                }
                if (loaded.Presentations.Count > 0)
                {
                    if (!BuildConfiguredPresentations(loaded, stageRoot, out var presentationError))
                    {
                        PrintDiagnostics(presentationError, loaded.Root);
                        fail(); return;
                    }
                    Console.WriteLine($"build: wrote {loaded.Presentations.Count} presentation(s) to {Path.Combine(loaded.Root, "YadgPresentations")}");
                }
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
        var verbose = new Option<bool>("--verbose") { Description = "Write invocation-local renderer phase timings to stderr." };
        command.Add(workspace); command.Add(renderer); command.Add(rendererPath); command.Add(yolo); command.Add(verbose);
        command.SetAction(parseResult =>
        {
            var path = parseResult.GetValue(workspace);
            var rendererId = parseResult.GetValue(renderer)!;
            var executable = parseResult.GetValue(rendererPath);
            var useYolo = parseResult.GetValue(yolo);
            var useVerbose = parseResult.GetValue(verbose);
            RenderResult result;
            if (!string.Equals(rendererId, "word", StringComparison.OrdinalIgnoreCase) && !string.Equals(rendererId, "libreoffice", StringComparison.OrdinalIgnoreCase)) { Console.Error.WriteLine($"YADG-RENDER-020: Unsupported renderer ID '{rendererId}'. Supported renderers are 'libreoffice' and 'word'."); fail(); return; }
            if (string.Equals(rendererId, "word", StringComparison.OrdinalIgnoreCase) && executable is not null) { Console.Error.WriteLine("YADG-RENDER-021: --renderer-path is valid only for LibreOffice."); fail(); return; }
            var root = Path.GetFullPath(path?.FullName ?? Directory.GetCurrentDirectory());
            if (useYolo)
            {
                result = RunStagedRenderer(rendererId, root, executable?.FullName, useVerbose);
                if (!result.Success)
                {
                    var alternate = rendererId.Equals("word", StringComparison.OrdinalIgnoreCase) ? "libreoffice" : "word";
                    var second = RunStagedRenderer(alternate, root, null, useVerbose);
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
            if (string.Equals(rendererId, "word", StringComparison.OrdinalIgnoreCase))
            {
                var clock = Stopwatch.StartNew();
                if (useVerbose) Console.Error.WriteLine("[renderer] phase started: Microsoft Word render");
                result = new WordRendererEngine().Render(root);
                if (useVerbose) Console.Error.WriteLine($"[renderer] phase {(result.Success ? "complete" : "failed")}: Microsoft Word render; elapsed={clock.Elapsed.TotalMilliseconds:F0}ms");
            }
            else result = new LibreOfficeRenderer().Render(root, executable?.FullName, useVerbose ? Console.Error.WriteLine : null);
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

    private static Command CreatePresentationCommand(Action fail)
    {
        var command = new Command("presentation", "Validate or build an editable PowerPoint presentation without workspace configuration.");
        var check = new Command("check", "Validate explicit Markdown sources against one PowerPoint template.");
        var checkTemplate = RepeatedOption("--template", "PowerPoint .pptx template path.");
        var checkSource = RepeatedOption("--source", "Markdown slide source path; repeat to preserve source order.");
        var checkOutput = RepeatedOption("--output", "Not valid with presentation check.");
        check.Add(checkTemplate); check.Add(checkSource); check.Add(checkOutput);
        check.SetAction(result =>
        {
            var templates = result.GetValue(checkTemplate) ?? Array.Empty<string>();
            var sources = result.GetValue(checkSource) ?? Array.Empty<string>();
            var output = result.GetValue(checkOutput) ?? Array.Empty<string>();
            if (output.Length > 0) { Console.Error.WriteLine("error YADG-PPT-CLI-001: presentation check does not accept --output."); fail(); return; }
            var analysis = AnalyzeDirectPresentation(templates, sources);
            PrintDiagnostics(analysis.Diagnostics);
            if (!analysis.IsValid) { Console.WriteLine("presentation check: invalid"); fail(); return; }
            Console.WriteLine($"presentation check: valid ({analysis.Document.Slides.Count} slide(s))");
        });

        var build = new Command("build", "Build one editable PowerPoint deck from explicit Markdown sources.");
        var buildTemplate = RepeatedOption("--template", "PowerPoint .pptx template path.");
        var buildSource = RepeatedOption("--source", "Markdown slide source path; repeat to preserve source order.");
        var buildOutput = RepeatedOption("--output", "Explicit output .pptx path.");
        build.Add(buildTemplate); build.Add(buildSource); build.Add(buildOutput);
        build.SetAction(result =>
        {
            var templates = result.GetValue(buildTemplate) ?? Array.Empty<string>();
            var sources = result.GetValue(buildSource) ?? Array.Empty<string>();
            var outputs = result.GetValue(buildOutput) ?? Array.Empty<string>();
            var analysis = AnalyzeDirectPresentation(templates, sources);
            if (!analysis.IsValid)
            {
                PrintDiagnostics(analysis.Diagnostics);
                Console.WriteLine("presentation build: invalid");
                fail(); return;
            }
            if (outputs.Length != 1) { Console.Error.WriteLine("error YADG-PPT-CLI-002: presentation build requires exactly one --output."); fail(); return; }
            if (!TryResolveDirectOutput(outputs[0], templates[0], sources, out var outputPath, out var outputError))
            { Console.Error.WriteLine($"error YADG-PPT-CLI-003: {outputError}"); fail(); return; }
            var directory = Path.GetDirectoryName(outputPath)!;
            var temporary = Path.Combine(directory, $".{Path.GetFileName(outputPath)}.yadg-stage-{Guid.NewGuid():N}.pptx");
            try
            {
                Directory.CreateDirectory(directory);
                var sourcesLoaded = ReadDirectSources(sources, out var sourceError);
                if (sourcesLoaded is null) { Console.Error.WriteLine($"error YADG-PPT-CLI-004: {sourceError}"); fail(); return; }
                var diagnostics = PresentationAuthoring.Build(Path.GetFullPath(templates[0]), sourcesLoaded, Directory.GetCurrentDirectory(), temporary);
                PrintDiagnostics(diagnostics);
                if (diagnostics.Any(diagnostic => diagnostic.IsError)) { fail(); return; }
                File.Move(temporary, outputPath, true);
                Console.WriteLine($"presentation build: wrote {outputPath}");
            }
            catch (Exception ex) { Console.Error.WriteLine($"error YADG-PPT-CLI-005: Unable to build presentation: {ex.Message}"); fail(); }
            finally { if (File.Exists(temporary)) try { File.Delete(temporary); } catch { } }
        });

        command.Add(check); command.Add(build);
        return command;
    }

    private static Option<string[]?> RepeatedOption(string name, string description) =>
        new(name) { Description = description, AllowMultipleArgumentsPerToken = false };

    private static PresentationAnalysis AnalyzeDirectPresentation(string[] templates, string[] sources)
    {
        var diagnostics = new List<Diagnostic>();
        if (templates.Length != 1) diagnostics.Add(new("YADG-PPT-CLI-001", "Presentation command requires exactly one --template."));
        if (sources.Length == 0) diagnostics.Add(new("YADG-PPT-CLI-001", "Presentation command requires one or more --source arguments."));
        if (diagnostics.Count > 0) return new(new(Array.Empty<PresentationSlide>(), diagnostics), diagnostics);
        if (!ValidateDirectInputs(templates[0], sources, out var templatePath, out var sourcePaths, out var error))
        {
            diagnostics.Add(new("YADG-PPT-CLI-004", error!));
            return new(new(Array.Empty<PresentationSlide>(), diagnostics), diagnostics);
        }
        var content = ReadDirectSources(sourcePaths, out error);
        if (content is null)
        {
            diagnostics.Add(new("YADG-PPT-CLI-004", error!));
            return new(new(Array.Empty<PresentationSlide>(), diagnostics), diagnostics);
        }
        return PresentationAuthoring.Analyze(templatePath, content, Directory.GetCurrentDirectory());
    }

    private static bool ValidateDirectInputs(string template, string[] sources, out string templatePath, out string[] sourcePaths, out string? error)
    {
        templatePath = string.Empty; sourcePaths = Array.Empty<string>(); error = null;
        if (IsRemoteOrPattern(template) || Directory.Exists(template) || !Path.GetExtension(template).Equals(".pptx", StringComparison.OrdinalIgnoreCase)) { error = "--template must identify one local .pptx file path."; return false; }
        templatePath = Path.GetFullPath(template);
        if (!File.Exists(templatePath)) { error = $"PowerPoint template does not exist: '{template}'."; return false; }
        var resolved = new List<string>();
        foreach (var source in sources)
        {
            if (IsRemoteOrPattern(source) || Directory.Exists(source) || !Path.GetExtension(source).Equals(".md", StringComparison.OrdinalIgnoreCase)) { error = $"--source must identify a local .md file path: '{source}'."; return false; }
            var full = Path.GetFullPath(source);
            if (!File.Exists(full)) { error = $"Markdown source does not exist: '{source}'."; return false; }
            resolved.Add(full);
        }
        sourcePaths = resolved.ToArray();
        return true;
    }

    private static IReadOnlyList<PresentationSource>? ReadDirectSources(string[] paths, out string? error)
    {
        error = null;
        try { return paths.Select(path => new PresentationSource(Path.GetFullPath(path), File.ReadAllText(path))).ToArray(); }
        catch (Exception ex) { error = ex.Message; return null; }
    }

    private static bool TryResolveDirectOutput(string output, string template, string[] sources, out string outputPath, out string? error)
    {
        outputPath = string.Empty; error = null;
        if (IsRemoteOrPattern(output) || Directory.Exists(output) || !Path.GetExtension(output).Equals(".pptx", StringComparison.OrdinalIgnoreCase)) { error = "--output must be a local .pptx file path."; return false; }
        outputPath = Path.GetFullPath(output);
        var resolvedOutput = outputPath;
        if (new[] { template }.Concat(sources).Any(path => string.Equals(Path.GetFullPath(path), resolvedOutput, StringComparison.OrdinalIgnoreCase))) { error = "--output must not overwrite the template or a Markdown source."; return false; }
        return true;
    }

    private static bool IsRemoteOrPattern(string path) => Uri.TryCreate(path, UriKind.Absolute, out var uri) && !uri.IsFile || path.Contains('*') || path.Contains('?');

    private static IReadOnlyList<Diagnostic> ValidatePresentations(YadgWorkspace workspace)
    {
        var diagnostics = new List<Diagnostic>();
        foreach (var deck in workspace.Presentations)
        {
            try
            {
                var sources = deck.SourcePaths.Select(path => new PresentationSource(path, File.ReadAllText(path))).ToArray();
                var analysis = PresentationAuthoring.Analyze(deck.TemplatePath, sources, workspace.Root, workspace.Values.CodeInlinePolicy);
                diagnostics.AddRange(analysis.Diagnostics.Select(diagnostic => diagnostic with { Message = $"Presentation '{deck.Declaration.Id}': {diagnostic.Message}" }));
            }
            catch (Exception ex) { diagnostics.Add(new("YADG-PPT-WORKSPACE-001", $"Presentation '{deck.Declaration.Id}' cannot be checked: {ex.Message}", true, deck.TemplatePath)); }
        }
        return diagnostics;
    }

    private static bool BuildConfiguredPresentations(YadgWorkspace workspace, string stageRoot, out IReadOnlyList<Diagnostic> diagnostics)
    {
        var allDiagnostics = new List<Diagnostic>();
        var loaded = new List<(WorkspaceLoader.ResolvedPresentation Deck, IReadOnlyList<PresentationSource> Sources)>();
        foreach (var deck in workspace.Presentations)
        {
            try
            {
                var sources = deck.SourcePaths.Select(path => new PresentationSource(path, File.ReadAllText(path))).ToArray();
                var analysis = PresentationAuthoring.Analyze(deck.TemplatePath, sources, workspace.Root, workspace.Values.CodeInlinePolicy);
                allDiagnostics.AddRange(analysis.Diagnostics.Select(diagnostic => diagnostic with { Message = $"Presentation '{deck.Declaration.Id}': {diagnostic.Message}" }));
                loaded.Add((deck, sources));
            }
            catch (Exception ex) { allDiagnostics.Add(new("YADG-PPT-WORKSPACE-001", $"Presentation '{deck.Declaration.Id}' cannot be checked: {ex.Message}", true, deck.TemplatePath)); }
        }
        if (allDiagnostics.Any(diagnostic => diagnostic.IsError)) { diagnostics = allDiagnostics; return false; }

        var staged = Path.Combine(stageRoot, "YadgPresentations");
        Directory.CreateDirectory(staged);
        try
        {
            foreach (var (deck, sources) in loaded)
            {
                var target = Path.Combine(staged, deck.Declaration.Output);
                allDiagnostics.AddRange(PresentationAuthoring.Build(deck.TemplatePath, sources, workspace.Root, target, workspace.Values.CodeInlinePolicy)
                    .Select(diagnostic => diagnostic with { Message = $"Presentation '{deck.Declaration.Id}': {diagnostic.Message}" }));
                if (allDiagnostics.Any(diagnostic => diagnostic.IsError)) { diagnostics = allDiagnostics; return false; }
            }
            CommitPresentationOutputs(staged, Path.Combine(workspace.Root, "YadgPresentations"));
        }
        catch (Exception ex) { allDiagnostics.Add(new("YADG-PPT-BUILD-001", $"Unable to stage or commit the presentation set: {ex.Message}", true, workspace.Root)); diagnostics = allDiagnostics; return false; }
        diagnostics = allDiagnostics;
        return true;
    }

    private static void CommitPresentationOutputs(string staged, string outputDirectory)
    {
        Directory.CreateDirectory(outputDirectory);
        var backup = Path.Combine(Path.GetDirectoryName(staged)!, "previous-presentations");
        Directory.CreateDirectory(backup);
        var moved = new List<(string Original, string Backup)>();
        var committed = new List<string>();
        try
        {
            foreach (var existing in Directory.EnumerateFiles(outputDirectory, "*.pptx", SearchOption.TopDirectoryOnly).ToArray())
            {
                var saved = Path.Combine(backup, Path.GetFileName(existing));
                File.Move(existing, saved, true);
                moved.Add((existing, saved));
            }
            foreach (var stagedFile in Directory.EnumerateFiles(staged, "*.pptx", SearchOption.TopDirectoryOnly))
            {
                var destination = Path.Combine(outputDirectory, Path.GetFileName(stagedFile));
                File.Move(stagedFile, destination, true);
                committed.Add(destination);
            }
        }
        catch
        {
            foreach (var file in committed) if (File.Exists(file)) File.Delete(file);
            foreach (var (original, saved) in moved) if (File.Exists(saved)) File.Move(saved, original, true);
            throw;
        }
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

    private static RenderResult RunStagedRenderer(string rendererId, string workspaceRoot, string? rendererPath, bool verbose)
    {
        var stage = Path.Combine(workspaceRoot, ".yadg-yolo-stage-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(Path.Combine(stage, "YadgPreWords"));
            foreach (var input in Directory.EnumerateFiles(Path.Combine(workspaceRoot, "YadgPreWords"), "*.docx", SearchOption.TopDirectoryOnly)) File.Copy(input, Path.Combine(stage, "YadgPreWords", Path.GetFileName(input)), true);
            RenderResult result;
            if (rendererId.Equals("word", StringComparison.OrdinalIgnoreCase))
            {
                var clock = Stopwatch.StartNew();
                if (verbose) Console.Error.WriteLine("[renderer] phase started: Microsoft Word staged render");
                result = new WordRendererEngine().Render(stage);
                if (verbose) Console.Error.WriteLine($"[renderer] phase {(result.Success ? "complete" : "failed")}: Microsoft Word staged render; elapsed={clock.Elapsed.TotalMilliseconds:F0}ms");
            }
            else result = new LibreOfficeRenderer().Render(stage, rendererPath, verbose ? message => Console.Error.WriteLine(message) : null);
            if (!result.Success) return result;
            FinalizedOutputCommit.Commit(workspaceRoot, Path.Combine(stage, "YadgWords"), result.ActualRenderer == "libreoffice" ? Path.Combine(stage, "YadgPdfs") : null);
            return result;
        }
        catch (Exception ex) { return new(false, new[] { new RendererDiagnostic("YADG-YOLO-STAGE-001", $"Staged renderer output could not be committed safely: {ex.Message}") }, ActualRenderer: rendererId); }
        finally { try { if (Directory.Exists(stage)) Directory.Delete(stage, true); } catch { } }
    }
}
