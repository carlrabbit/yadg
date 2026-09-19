using System.Diagnostics;
using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text.Json;

namespace Yadg.Renderer;

public sealed record RendererDiagnostic(string Code, string Message, string? Location = null)
{
    public override string ToString() => Location is null ? $"{Code}: {Message}" : $"{Code}: {Location}: {Message}";
}

public sealed record RenderResult(bool Success, IReadOnlyList<RendererDiagnostic> Diagnostics, string? RuntimeVersion = null, string? Executable = null, string? ProfileIdentity = null);

public sealed class LibreOfficeRenderer
{
    private const int TimeoutMs = 30000;

    public RenderResult Render(string workspaceRoot, string? rendererPath = null)
    {
        var diagnostics = new List<RendererDiagnostic>();
        var root = Path.GetFullPath(workspaceRoot);
        if (!Directory.Exists(root)) return Fail(diagnostics, "YADG-RENDER-001", $"Workspace directory does not exist: '{root}'.");
        var preWords = Path.Combine(root, "YadgPreWords");
        if (!Directory.Exists(preWords)) return Fail(diagnostics, "YADG-RENDER-002", $"Missing authored-input directory '{preWords}'.");
        var inputs = Directory.EnumerateFiles(preWords, "*.docx", SearchOption.TopDirectoryOnly).OrderBy(p => p, StringComparer.Ordinal).ToArray();
        if (inputs.Length == 0) return Fail(diagnostics, "YADG-RENDER-003", $"No top-level DOCX inputs were found in '{preWords}'.");
        var executable = ResolveExecutable(rendererPath, diagnostics);
        if (executable is null) return new(false, diagnostics);
        var version = ReadVersion(executable, diagnostics);
        if (version is null) return new(false, diagnostics, null, executable);
        var python = ResolvePython(executable, diagnostics);
        if (python is null) return new(false, diagnostics, version, executable);
        var sidecar = Path.Combine(AppContext.BaseDirectory, "uno_renderer.py");
        if (!File.Exists(sidecar)) return Fail(diagnostics, "YADG-RENDER-006", $"Renderer UNO sidecar is missing: '{sidecar}'.");

        var words = Path.Combine(root, "YadgWords"); var pdfs = Path.Combine(root, "YadgPdfs");
        var sessionId = Guid.NewGuid().ToString("N"); var sessionRoot = Path.Combine(Path.GetTempPath(), "yadg-lo-" + sessionId); var profile = Path.Combine(sessionRoot, "profile"); var stage = Path.Combine(sessionRoot, "stage");
        Directory.CreateDirectory(profile); Directory.CreateDirectory(stage);
        var port = GetFreePort(); Process? office = null;
        try
        {
            office = StartOffice(executable, profile, port, diagnostics);
            if (office is null) return new(false, diagnostics, version, executable, sessionId);
            foreach (var input in inputs)
            {
                var stagedInput = Path.Combine(stage, "input-" + Path.GetFileName(input));
                NormalizeImageRelationships(input, stagedInput);
                var stagedDocx = Path.Combine(stage, Path.GetFileName(input)); var stagedPdf = Path.Combine(stage, Path.GetFileNameWithoutExtension(input) + ".pdf");
                var sidecarResult = RunSidecar(python, sidecar, port, stagedInput, stagedDocx, stagedPdf, diagnostics);
                if (!sidecarResult) return new(false, diagnostics, version, executable, sessionId);
                if (!File.Exists(stagedDocx) || new FileInfo(stagedDocx).Length == 0) return Fail(diagnostics, "YADG-RENDER-014", $"LibreOffice did not produce a non-empty DOCX for '{input}'.", version, executable, sessionId);
                if (!File.Exists(stagedPdf) || new FileInfo(stagedPdf).Length == 0) return Fail(diagnostics, "YADG-RENDER-015", $"LibreOffice did not produce a non-empty PDF for '{input}'.", version, executable, sessionId);
            }
            Directory.CreateDirectory(words); Directory.CreateDirectory(pdfs);
            foreach (var input in inputs)
            {
                File.Move(Path.Combine(stage, Path.GetFileName(input)), Path.Combine(words, Path.GetFileName(input)), true);
                File.Move(Path.Combine(stage, Path.GetFileNameWithoutExtension(input) + ".pdf"), Path.Combine(pdfs, Path.GetFileNameWithoutExtension(input) + ".pdf"), true);
            }
            return new(true, diagnostics, version, executable, sessionId);
        }
        catch (Exception ex) { return Fail(diagnostics, "YADG-RENDER-016", $"LibreOffice rendering failed: {ex.Message}", version, executable, sessionId); }
        finally
        {
            try { if (office is not null && !office.HasExited) office.Kill(true); } catch { }
            try { office?.Dispose(); } catch { }
            try { if (Directory.Exists(sessionRoot)) Directory.Delete(sessionRoot, true); } catch { }
        }
    }

    private static Process? StartOffice(string executable, string profile, int port, List<RendererDiagnostic> diagnostics)
    {
        var profileUrl = new Uri(profile + Path.DirectorySeparatorChar).AbsoluteUri;
        var info = new ProcessStartInfo(executable, $"--headless --norestore --nofirststartwizard --nodefault --nolockcheck --accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext -env:UserInstallation={profileUrl}") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true };
        var process = Process.Start(info); if (process is null) { diagnostics.Add(new("YADG-RENDER-007", "Unable to start LibreOffice.")); return null; }
        return process;
    }

    private static bool RunSidecar(string python, string sidecar, int port, string input, string docx, string pdf, List<RendererDiagnostic> diagnostics)
    {
        var info = new ProcessStartInfo(python) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true };
        info.ArgumentList.Add(sidecar); info.ArgumentList.Add("--port"); info.ArgumentList.Add(port.ToString()); info.ArgumentList.Add("--input"); info.ArgumentList.Add(input); info.ArgumentList.Add("--docx-output"); info.ArgumentList.Add(docx); info.ArgumentList.Add("--pdf-output"); info.ArgumentList.Add(pdf);
        using var process = Process.Start(info); if (process is null) { diagnostics.Add(new("YADG-RENDER-008", "Unable to start the LibreOffice UNO bridge.")); return false; }
        if (!process.WaitForExit(TimeoutMs)) { try { process.Kill(true); } catch { } diagnostics.Add(new("YADG-RENDER-009", "Timed out waiting for the LibreOffice UNO operation.")); return false; }
        var stdout = process.StandardOutput.ReadToEnd(); var stderr = process.StandardError.ReadToEnd();
        if (process.ExitCode != 0) { diagnostics.Add(new("YADG-RENDER-010", $"LibreOffice UNO operation failed for '{input}': {stdout} {stderr}")); return false; }
        try { using var json = JsonDocument.Parse(stdout); if (!json.RootElement.GetProperty("ok").GetBoolean()) { diagnostics.Add(new("YADG-RENDER-011", json.RootElement.GetProperty("error").GetString() ?? "LibreOffice reported an unspecified failure.", input)); return false; } }
        catch (Exception ex) { diagnostics.Add(new("YADG-RENDER-012", $"LibreOffice UNO bridge returned invalid status: {ex.Message}", input)); return false; }
        return true;
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
        try { using var process = Process.Start(new ProcessStartInfo(executable, "--version") { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true }); if (process is null || !process.WaitForExit(10000)) { diagnostics.Add(new("YADG-RENDER-017", "Timed out obtaining the LibreOffice runtime version.")); return null; } var text = (process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd()).Trim(); if (process.ExitCode != 0 || text.Length == 0) { diagnostics.Add(new("YADG-RENDER-018", "LibreOffice version detection failed.")); return null; } return text; }
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
                // canonical relative target after moving the package part.
                var updated = xml.Replace("/" + source.FullName, "/" + targetName, StringComparison.Ordinal);
                if (!string.Equals(xml, updated, StringComparison.Ordinal)) replacements.Add((relationship, updated));
            }
            source.Delete();
            foreach (var (relationship, content) in replacements) { var name = relationship.FullName; relationship.Delete(); var replacement = archive.CreateEntry(name, CompressionLevel.Optimal); using var write = new StreamWriter(replacement.Open(), new System.Text.UTF8Encoding(false)); write.Write(content); }
        }
    }
    private static RenderResult Fail(List<RendererDiagnostic> diagnostics, string code, string message, string? version = null, string? executable = null, string? profile = null) { diagnostics.Add(new(code, message)); return new(false, diagnostics, version, executable, profile); }
}
