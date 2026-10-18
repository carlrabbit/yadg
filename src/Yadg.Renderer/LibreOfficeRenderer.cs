using System.Diagnostics;
using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Yadg.Renderer;

public sealed record RendererDiagnostic(string Code, string Message, string? Location = null)
{
    public override string ToString() => Location is null ? $"{Code}: {Message}" : $"{Code}: {Location}: {Message}";
}

public sealed record RenderResult(bool Success, IReadOnlyList<RendererDiagnostic> Diagnostics, string? RuntimeVersion = null, string? Executable = null, string? ProfileIdentity = null, string ActualRenderer = "unknown");

public sealed class LibreOfficeRenderer
{
    private const int TimeoutMs = 30000;

    public RenderResult Render(string workspaceRoot, string? rendererPath = null, Action<string>? verbose = null)
    {
        var renderClock = Stopwatch.StartNew();
        void Log(string message) => verbose?.Invoke($"[renderer +{renderClock.Elapsed.TotalSeconds:F3}s] {message}");
        void StartPhase(string phase, string? input = null) => Log($"phase started: {phase}" + (input is null ? string.Empty : $"; input={input}"));
        void Timed(string phase, Stopwatch clock, string? input = null) => Log($"phase complete: {phase}; elapsed={clock.Elapsed.TotalMilliseconds:F0}ms" + (input is null ? string.Empty : $"; input={input}"));
        RenderResult FailWithTotal(List<RendererDiagnostic> items, string code, string message, string? version = null, string? executable = null, string? profile = null)
        {
            Timed("total render", renderClock);
            return Fail(items, code, message, version, executable, profile);
        }
        var diagnostics = new List<RendererDiagnostic>();
        var root = Path.GetFullPath(workspaceRoot);
        if (!Directory.Exists(root)) return FailWithTotal(diagnostics, "YADG-RENDER-001", $"Workspace directory does not exist: '{root}'.");
        var preWords = Path.Combine(root, "YadgPreWords");
        if (!Directory.Exists(preWords)) return FailWithTotal(diagnostics, "YADG-RENDER-002", $"Missing authored-input directory '{preWords}'.");
        var inputs = Directory.EnumerateFiles(preWords, "*.docx", SearchOption.TopDirectoryOnly).OrderBy(p => p, StringComparer.Ordinal).ToArray();
        if (inputs.Length == 0) return FailWithTotal(diagnostics, "YADG-RENDER-003", $"No top-level DOCX inputs were found in '{preWords}'.");
        var phaseClock = Stopwatch.StartNew();
        StartPhase("runtime/executable resolution");
        var executable = ResolveExecutable(rendererPath, diagnostics);
        Timed("runtime/executable resolution", phaseClock);
        if (executable is null) { Timed("total render", renderClock); return new(false, diagnostics, ActualRenderer: "libreoffice"); }
        phaseClock.Restart();
        StartPhase("version/runtime discovery");
        var version = ReadVersion(executable, diagnostics);
        Timed("version/runtime discovery", phaseClock);
        if (version is null) { Timed("total render", renderClock); return new(false, diagnostics, null, executable, ActualRenderer: "libreoffice"); }
        phaseClock.Restart();
        StartPhase("UNO runtime resolution");
        var python = ResolvePython(executable, diagnostics);
        Timed("UNO runtime resolution", phaseClock);
        if (python is null) { Timed("total render", renderClock); return new(false, diagnostics, version, executable, ActualRenderer: "libreoffice"); }
        phaseClock.Restart();
        StartPhase("UNO sidecar resolution");
        var sidecar = Path.Combine(AppContext.BaseDirectory, "uno_renderer.py");
        if (!File.Exists(sidecar)) return FailWithTotal(diagnostics, "YADG-RENDER-006", $"Renderer UNO sidecar is missing: '{sidecar}'.");
        Timed("UNO sidecar resolution", phaseClock);

        var words = Path.Combine(root, "YadgWords"); var pdfs = Path.Combine(root, "YadgPdfs");
        var sessionId = Guid.NewGuid().ToString("N"); var sessionRoot = Path.Combine(Path.GetTempPath(), "yadg-lo-" + sessionId); var profile = Path.Combine(sessionRoot, "profile"); var stage = Path.Combine(sessionRoot, "stage");
        Directory.CreateDirectory(profile); Directory.CreateDirectory(stage);
        var port = GetFreePort(); Process? office = null;
        Task officeOutputTask = Task.CompletedTask;
        try
        {
            phaseClock.Restart();
            StartPhase($"isolated session/process start; session={sessionId}");
            office = StartOffice(executable, profile, port, diagnostics);
            if (office is null) return new(false, diagnostics, version, executable, sessionId, "libreoffice");
            officeOutputTask = DrainChildOutputAsync(office.StandardOutput, office.StandardError, null, null);
            Timed($"isolated session/process start; session={sessionId}", phaseClock);
            foreach (var input in inputs)
            {
                var inputName = Path.GetFileName(input);
                var inputClock = Stopwatch.StartNew();
                var stagedInput = Path.Combine(stage, "input-" + Path.GetFileName(input));
                phaseClock.Restart();
                StartPhase("per-document package normalization", inputName);
                NormalizeImageRelationships(input, stagedInput);
                Timed("per-document package normalization", phaseClock, inputName);
                var stagedDocx = Path.Combine(stage, Path.GetFileName(input)); var stagedPdf = Path.Combine(stage, Path.GetFileNameWithoutExtension(input) + ".pdf");
                var sidecarResult = RunSidecar(python, sidecar, port, stagedInput, stagedDocx, stagedPdf, diagnostics, verbose is null ? null : Log, inputName);
                if (!sidecarResult) { Timed("per-document total", inputClock, inputName); return new(false, diagnostics, version, executable, sessionId, "libreoffice"); }
                phaseClock.Restart();
                StartPhase("output validation", inputName);
                if (!File.Exists(stagedDocx) || new FileInfo(stagedDocx).Length == 0) { Timed("per-document total", inputClock, inputName); return Fail(diagnostics, "YADG-RENDER-014", $"LibreOffice did not produce a non-empty DOCX for '{input}'.", version, executable, sessionId); }
                if (!File.Exists(stagedPdf) || new FileInfo(stagedPdf).Length == 0) { Timed("per-document total", inputClock, inputName); return Fail(diagnostics, "YADG-RENDER-015", $"LibreOffice did not produce a non-empty PDF for '{input}'.", version, executable, sessionId); }
                Timed("output validation", phaseClock, inputName);
                Timed("per-document total", inputClock, inputName);
            }
            var stagedWords = Path.Combine(stage, "YadgWords"); var stagedPdfs = Path.Combine(stage, "YadgPdfs");
            Directory.CreateDirectory(stagedWords); Directory.CreateDirectory(stagedPdfs);
            phaseClock.Restart();
            StartPhase("output validation/commit");
            foreach (var input in inputs)
            {
                File.Move(Path.Combine(stage, Path.GetFileName(input)), Path.Combine(stagedWords, Path.GetFileName(input)));
                File.Move(Path.Combine(stage, Path.GetFileNameWithoutExtension(input) + ".pdf"), Path.Combine(stagedPdfs, Path.GetFileNameWithoutExtension(input) + ".pdf"));
            }
            FinalizedOutputCommit.Commit(root, stagedWords, stagedPdfs);
            Timed("output validation/commit", phaseClock);
            return new(true, diagnostics, version, executable, sessionId, "libreoffice");
        }
        catch (Exception ex) { return Fail(diagnostics, "YADG-RENDER-016", $"LibreOffice rendering failed: {ex.Message}", version, executable, sessionId); }
        finally
        {
            phaseClock.Restart();
            try { if (office is not null && !office.HasExited) office.Kill(true); } catch { }
            try { office?.Dispose(); } catch { }
            try { officeOutputTask.GetAwaiter().GetResult(); } catch { }
            try { if (Directory.Exists(sessionRoot)) Directory.Delete(sessionRoot, true); } catch { }
            Timed("cleanup", phaseClock);
            Timed("total render", renderClock);
        }
    }

    private static Process? StartOffice(string executable, string profile, int port, List<RendererDiagnostic> diagnostics)
    {
        var profileUrl = new Uri(profile + Path.DirectorySeparatorChar).AbsoluteUri;
        var info = new ProcessStartInfo(executable, $"--headless --norestore --nofirststartwizard --nodefault --nolockcheck --accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext -env:UserInstallation={profileUrl}") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true };
        var process = Process.Start(info); if (process is null) { diagnostics.Add(new("YADG-RENDER-007", "Unable to start LibreOffice.")); return null; }
        return process;
    }

    internal static bool RunSidecar(string python, string sidecar, int port, string input, string docx, string pdf, List<RendererDiagnostic> diagnostics, Action<string>? verbose, string inputIdentity, int timeoutMs = TimeoutMs)
    {
        var info = new ProcessStartInfo(python) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true };
        info.ArgumentList.Add(sidecar); info.ArgumentList.Add("--port"); info.ArgumentList.Add(port.ToString()); info.ArgumentList.Add("--input"); info.ArgumentList.Add(input); info.ArgumentList.Add("--docx-output"); info.ArgumentList.Add(docx); info.ArgumentList.Add("--pdf-output"); info.ArgumentList.Add(pdf);
        using var process = Process.Start(info); if (process is null) { diagnostics.Add(new("YADG-RENDER-008", "Unable to start the LibreOffice UNO bridge.")); return false; }
        var clock = Stopwatch.StartNew();
        var currentPhase = "UNO connect";
        var lastCompletedPhase = "none";
        var phaseLock = new object();
        var stdout = new StringBuilder();
        var stderr = new StringBuilder();
        var outputLock = new object();
        var stdoutTask = DrainSidecarStdoutAsync(process.StandardOutput, stdout, outputLock, phase =>
        {
            lock (phaseLock)
            {
                currentPhase = phase.Current;
                if (phase.Completed is not null) lastCompletedPhase = phase.Completed;
            }
            if (phase.Message is not null) verbose?.Invoke($"[LibreOffice input={inputIdentity} +{clock.Elapsed.TotalSeconds:F3}s] {phase.Message}");
        });
        var stderrTask = CaptureChildOutputAsync(process.StandardError, stderr, outputLock);
        var waitTask = process.WaitForExitAsync();
        if (Task.WhenAny(waitTask, Task.Delay(timeoutMs)).GetAwaiter().GetResult() != waitTask)
        {
            var elapsed = clock.Elapsed;
            string timeoutPhase; string timeoutLastCompleted;
            lock (phaseLock) { timeoutPhase = currentPhase; timeoutLastCompleted = lastCompletedPhase; }
            try { process.Kill(true); } catch { }
            try { waitTask.GetAwaiter().GetResult(); } catch { }
            try { Task.WhenAll(stdoutTask, stderrTask).GetAwaiter().GetResult(); } catch { }
            diagnostics.Add(new("YADG-RENDER-009", $"Timed out waiting for the LibreOffice UNO operation after {timeoutMs} ms (timeout={timeoutMs} ms; elapsed={elapsed.TotalMilliseconds:F0} ms; input='{inputIdentity}'; current phase='{timeoutPhase}'; last completed phase='{timeoutLastCompleted}')."));
            return false;
        }
        Task.WhenAll(stdoutTask, stderrTask).GetAwaiter().GetResult();
        string resultOutput; string stderrOutput;
        lock (outputLock) { resultOutput = stdout.ToString(); stderrOutput = stderr.ToString(); }
        var lines = resultOutput.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var resultLine = lines.LastOrDefault();
        var childOutput = string.Join(Environment.NewLine, lines.Take(lines.Length - (resultLine is null ? 0 : 1)));
        if (childOutput.Length > 0) verbose?.Invoke($"[LibreOffice input={inputIdentity}] sidecar output: {childOutput}");
        if (process.ExitCode != 0) { diagnostics.Add(new("YADG-RENDER-010", $"LibreOffice UNO operation failed for '{input}': {resultLine} {stderrOutput}")); return false; }
        try { using var json = JsonDocument.Parse(resultLine ?? string.Empty); if (!json.RootElement.GetProperty("ok").GetBoolean()) { diagnostics.Add(new("YADG-RENDER-011", json.RootElement.GetProperty("error").GetString() ?? "LibreOffice reported an unspecified failure.", input)); return false; } }
        catch (Exception ex) { diagnostics.Add(new("YADG-RENDER-012", $"LibreOffice UNO bridge returned invalid status: {ex.Message}", input)); return false; }
        return true;
    }

    private sealed record SidecarPhase(string Current, string? Completed, string? Message);

    private static async Task DrainSidecarStdoutAsync(StreamReader reader, StringBuilder output, object sync, Action<SidecarPhase> phaseObserved)
    {
        while (await reader.ReadLineAsync().ConfigureAwait(false) is { } line)
        {
            try
            {
                using var json = JsonDocument.Parse(line);
                var root = json.RootElement;
                if (root.TryGetProperty("type", out var type) && type.GetString() == "progress")
                {
                    var phase = root.GetProperty("phase").GetString() ?? "unknown";
                    var status = root.GetProperty("status").GetString();
                    var elapsed = root.TryGetProperty("elapsedMs", out var elapsedValue) ? elapsedValue.GetDouble() : 0d;
                    var duration = root.TryGetProperty("phaseMs", out var durationValue) ? durationValue.GetDouble() : (double?)null;
                    phaseObserved(new(status == "complete" ? $"after {phase}" : phase, status == "complete" ? phase : null,
                        status == "complete" ? $"phase complete: {phase}; phaseElapsed={duration.GetValueOrDefault():F0}ms; sidecarElapsed={elapsed:F0}ms" : $"phase started: {phase}; sidecarElapsed={elapsed:F0}ms"));
                    continue;
                }
            }
            catch (Exception) { }
            if (line.Length > 2048) line = line[..2048] + "…";
            lock (sync)
            {
                output.AppendLine(line);
                if (output.Length > 16_384) output.Remove(0, output.Length - 16_384);
            }
        }
    }

    private static async Task CaptureChildOutputAsync(StreamReader reader, StringBuilder destination, object sync)
    {
        while (await reader.ReadLineAsync().ConfigureAwait(false) is { } line)
        {
            // Keep diagnostics bounded while continuing to drain the pipe.
            if (line.Length > 2048) line = line[..2048] + "…";
            lock (sync)
            {
                destination.AppendLine(line);
                if (destination.Length > 16_384) destination.Remove(0, destination.Length - 16_384);
            }
        }
    }

    private static async Task DrainChildOutputAsync(StreamReader stdout, StreamReader stderr, Action<string>? onOutput, Action<string>? onError)
    {
        await Task.WhenAll(DrainOne(stdout, onOutput), DrainOne(stderr, onError)).ConfigureAwait(false);
        static async Task DrainOne(StreamReader reader, Action<string>? sink)
        {
            if (sink is not null)
            {
                while (await reader.ReadLineAsync().ConfigureAwait(false) is { } line) sink(line.Length > 2048 ? line[..2048] + "…" : line);
                return;
            }
            var buffer = new char[4096];
            while (await reader.ReadAsync(buffer.AsMemory()).ConfigureAwait(false) > 0) { }
        }
    }

    private static string? ResolveExecutable(string? requested, List<RendererDiagnostic> diagnostics)
    {
        if (!string.IsNullOrWhiteSpace(requested)) { var path = Path.GetFullPath(requested); if (File.Exists(path)) return path; diagnostics.Add(new("YADG-RENDER-004", $"Specified LibreOffice executable does not exist: '{path}'.")); return null; }
        var names = OperatingSystem.IsWindows() ? new[] { "soffice.com", "soffice.exe" } : new[] { "libreoffice", "soffice" };
        foreach (var directory in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)) foreach (var name in names) { var path = Path.Combine(directory, name); if (File.Exists(path)) return path; }
        diagnostics.Add(new("YADG-RENDER-005", "LibreOffice was not found on PATH. Provide --renderer-path with soffice.com/soffice.", null)); return null;
    }

    private static string? ResolvePython(string executable, List<RendererDiagnostic> diagnostics)
    {
        var sibling = Path.Combine(Path.GetDirectoryName(executable)!, "python.exe"); if (File.Exists(sibling)) return sibling;
        diagnostics.Add(new("YADG-RENDER-013", "The selected LibreOffice installation has no bundled Python/UNO runtime.")); return null;
    }

    private static string? ReadVersion(string executable, List<RendererDiagnostic> diagnostics)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo(executable, "--version") { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true });
            if (process is null) { diagnostics.Add(new("YADG-RENDER-018", "LibreOffice version detection failed.")); return null; }
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(10000))
            {
                try { process.Kill(true); } catch { }
                try { process.WaitForExit(); } catch { }
                try { Task.WhenAll(stdout, stderr).GetAwaiter().GetResult(); } catch { }
                diagnostics.Add(new("YADG-RENDER-017", "Timed out obtaining the LibreOffice runtime version."));
                return null;
            }
            var text = (stdout.GetAwaiter().GetResult() + stderr.GetAwaiter().GetResult()).Trim();
            if (process.ExitCode != 0 || text.Length == 0) { diagnostics.Add(new("YADG-RENDER-018", "LibreOffice version detection failed.")); return null; }
            return text;
        }
        catch (Exception ex) { diagnostics.Add(new("YADG-RENDER-018", $"LibreOffice version detection failed: {ex.Message}")); return null; }
    }

    private static int GetFreePort() { using var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start(); return ((IPEndPoint)listener.LocalEndpoint).Port; }
    private static void NormalizeImageRelationships(string input, string staged)
    {
        File.Copy(input, staged, true);
        using var archive = ZipFile.Open(staged, ZipArchiveMode.Update);
        var rootImages = archive.Entries.Where(e => e.FullName.StartsWith("media/", StringComparison.OrdinalIgnoreCase) && !e.FullName.EndsWith("/", StringComparison.Ordinal)).ToArray();
        foreach (var source in rootImages)
        {
            var targetName = "word/" + source.FullName;
            var target = archive.CreateEntry(targetName, CompressionLevel.Optimal);
            using (var sourceStream = source.Open()) using (var targetStream = target.Open()) sourceStream.CopyTo(targetStream);
            var replacements = new List<(ZipArchiveEntry Entry, string Content)>();
            foreach (var relationship in archive.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) || e.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)).ToArray())
            {
                string xml; using (var read = relationship.Open()) using (var buffer = new MemoryStream()) { read.CopyTo(buffer); xml = System.Text.Encoding.UTF8.GetString(buffer.ToArray()); }
                // A document relationship is resolved relative to /word/, so use the
                // canonical relative target after moving the package part. Content
                // types use package-root part names and must not receive this rewrite.
                var updated = relationship.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
                    ? xml.Replace("/" + source.FullName, targetName["word/".Length..], StringComparison.Ordinal)
                    : xml;
                if (!string.Equals(xml, updated, StringComparison.Ordinal)) replacements.Add((relationship, updated));
            }
            source.Delete();
            foreach (var (relationship, content) in replacements) { var name = relationship.FullName; relationship.Delete(); var replacement = archive.CreateEntry(name, CompressionLevel.Optimal); using var write = new StreamWriter(replacement.Open(), new System.Text.UTF8Encoding(false)); write.Write(content); }
        }
    }
    private static RenderResult Fail(List<RendererDiagnostic> diagnostics, string code, string message, string? version = null, string? executable = null, string? profile = null) { diagnostics.Add(new(code, message)); return new(false, diagnostics, version, executable, profile, "libreoffice"); }
}
