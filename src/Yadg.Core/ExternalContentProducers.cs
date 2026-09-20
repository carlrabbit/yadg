using System.Diagnostics;
using System.Text;

namespace Yadg.Core;

public static class MermaidProducer
{
    private const int TimeoutMilliseconds = 60_000;

    public static YadgDocument RenderFigures(YadgDocument document, string workspaceRoot, WorkspaceValues values, List<Diagnostic> diagnostics, List<string> temporaryPaths)
    {
        var replacements = new Dictionary<string, YadgFigure>(StringComparer.Ordinal);
        foreach (var figure in document.Figures.Values.Where(f => f.GeneratedSource is not null))
        {
            if (!values.Producers.TryGetValue("mermaid", out var config))
            {
                diagnostics.Add(new("YADG-PRODUCER-007", $"Mermaid figure '{figure.Id}' requires configured producers.mermaid.", true, figure.SourcePath));
                replacements[figure.Id] = figure;
                continue;
            }
            var rendered = RenderOne(figure, config, workspaceRoot, diagnostics);
            if (rendered is null) { replacements[figure.Id] = figure; continue; }
            replacements[figure.Id] = figure with { AssetPath = rendered.Value.Asset.FullPath, Asset = rendered.Value.Asset };
            temporaryPaths.Add(rendered.Value.Directory);
        }

        YadgBlock Map(YadgBlock block) => block is YadgFigure figure && replacements.TryGetValue(figure.Id, out var replacement) ? replacement : block;
        var blocks = document.Blocks.Select(Map).ToArray();
        var sections = document.References.ToDictionary(p => p.Key, p => p.Value with { Body = p.Value.Body.Select(Map).ToArray() }, StringComparer.Ordinal);
        var figures = document.Figures.ToDictionary(p => p.Key, p => replacements.TryGetValue(p.Key, out var replacement) ? replacement : p.Value, StringComparer.Ordinal);
        return document with { Blocks = blocks, References = sections, Figures = figures };
    }

    private static (ImageAsset Asset, string Directory)? RenderOne(YadgFigure figure, MermaidProducerConfiguration config, string workspaceRoot, List<Diagnostic> diagnostics)
    {
        var directory = Path.Combine(Path.GetTempPath(), "yadg-mermaid-" + Guid.NewGuid().ToString("N"));
        var input = Path.Combine(directory, "input.mmd");
        var output = Path.Combine(directory, "output.png");
        try
        {
            Directory.CreateDirectory(directory);
            File.WriteAllText(input, figure.GeneratedSource!, new UTF8Encoding(false));
            var executable = ResolveExecutable(config.Executable, workspaceRoot);
            var psi = new ProcessStartInfo { FileName = executable, WorkingDirectory = workspaceRoot, UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
            foreach (var argument in config.Arguments) psi.ArgumentList.Add(argument);
            psi.ArgumentList.Add("-i"); psi.ArgumentList.Add(input); psi.ArgumentList.Add("-o"); psi.ArgumentList.Add(output);
            using var process = Process.Start(psi);
            if (process is null) { Fail(diagnostics, "YADG-PRODUCER-010", "Unable to start Mermaid producer.", figure); return null; }
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(TimeoutMilliseconds))
            {
                try { process.Kill(true); } catch { }
                Fail(diagnostics, "YADG-PRODUCER-011", "Mermaid producer timed out after 60 seconds.", figure);
                return null;
            }
            var detail = Bound(stderr.GetAwaiter().GetResult());
            _ = Bound(stdout.GetAwaiter().GetResult());
            if (process.ExitCode != 0) { Fail(diagnostics, "YADG-PRODUCER-012", $"Mermaid producer exited with code {process.ExitCode}." + (detail.Length == 0 ? string.Empty : $" stderr: {detail}"), figure); return null; }
            var dimensions = ReadPng(output);
            if (dimensions is null) { Fail(diagnostics, "YADG-PRODUCER-013", "Mermaid producer did not create a valid non-empty PNG.", figure); return null; }
            return (new ImageAsset(output, "PNG", dimensions.Value.Width, dimensions.Value.Height), directory);
        }
        catch (Exception ex) { Fail(diagnostics, "YADG-PRODUCER-010", $"Unable to run Mermaid producer: {ex.Message}", figure); return null; }
    }

    private static string ResolveExecutable(string executable, string workspaceRoot)
    {
        if (Path.IsPathRooted(executable)) return executable;
        if (executable.Contains(Path.DirectorySeparatorChar) || executable.Contains(Path.AltDirectorySeparatorChar)) return Path.GetFullPath(Path.Combine(workspaceRoot, executable));
        return executable;
    }

    private static (int Width, int Height)? ReadPng(string path)
    {
        if (!File.Exists(path)) return null;
        var bytes = File.ReadAllBytes(path);
        if (bytes.Length < 33 || bytes[0] != 137 || bytes[1] != 80 || bytes[2] != 78 || bytes[3] != 71 || bytes[4] != 13 || bytes[5] != 10 || bytes[6] != 26 || bytes[7] != 10) return null;
        if (bytes[12] != 73 || bytes[13] != 72 || bytes[14] != 68 || bytes[15] != 82) return null;
        var width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
        var height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
        var position = 8; var hasIdat = false; var hasIend = false;
        while (position + 12 <= bytes.Length)
        {
            var length = (bytes[position] << 24) | (bytes[position + 1] << 16) | (bytes[position + 2] << 8) | bytes[position + 3];
            if (length < 0 || position + 12L + length > bytes.Length) return null;
            var type = Encoding.ASCII.GetString(bytes, position + 4, 4);
            if (type == "IDAT") hasIdat = true;
            if (type == "IEND") { hasIend = length == 0; break; }
            position += 12 + length;
        }
        return width > 0 && height > 0 && hasIdat && hasIend ? (width, height) : null;
    }

    private static string Bound(string text) => text.Length <= 4000 ? text.Trim() : text[..4000].Trim() + "…";
    private static void Fail(List<Diagnostic> diagnostics, string code, string message, YadgFigure figure) => diagnostics.Add(new(code, message, true, figure.SourcePath));
}
