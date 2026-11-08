using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Yadg.Core;

public sealed record WorkspaceValues(IReadOnlyDictionary<string, string> Values)
{
    public ReferenceFormatting References { get; init; } = ReferenceFormatting.Default;
    public IReadOnlyDictionary<string, MermaidProducerConfiguration> Producers { get; init; } = new Dictionary<string, MermaidProducerConfiguration>(StringComparer.Ordinal);
    public string? PublishPath { get; init; }
    public string ThematicBreakPolicy { get; init; } = "error";
    public string CodeInlinePolicy { get; init; } = "error";
    public bool DiagramsAutoDetect { get; init; }
    public IReadOnlyList<PresentationDeclaration> Presentations { get; init; } = Array.Empty<PresentationDeclaration>();
    public static WorkspaceValues Empty { get; } = new(new Dictionary<string, string>(StringComparer.Ordinal));
}

public sealed record MermaidProducerConfiguration(string Executable, IReadOnlyList<string> Arguments);
public sealed record PresentationDeclaration(string Id, string Template, IReadOnlyList<string> Sources, string Output);

public static class WorkspaceValuesParser
{
    private static readonly Regex Id = new("^[A-Za-z][A-Za-z0-9_-]*$", RegexOptions.Compiled);
    private static readonly Regex Number = new("^[+-]?(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?$", RegexOptions.Compiled);
    private static readonly Regex DateLike = new("^[0-9]{4}-[0-9]{2}-[0-9]{2}(?:[Tt ].*)?$", RegexOptions.Compiled);

    public static string? ParsePublishPath(string path, List<Diagnostic> diagnostics)
    {
        string text;
        try { text = File.ReadAllText(path, Encoding.UTF8); }
        catch (Exception ex) { diagnostics.Add(new("YADG-PUBLISH-009", $"Unable to read publication configuration: {ex.Message}", true, path)); return null; }
        if (text.Length > 0 && text[0] == '\ufeff') text = text[1..];
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var close = lines.Length > 0 && lines[0] == "---" ? Array.FindIndex(lines, 1, line => line == "---") : -1;
        if (close < 0) { diagnostics.Add(new("YADG-PUBLISH-009", "Cannot read publish.path: YADG.md has no valid front matter block.", true, path)); return null; }
        var inPublish = false;
        var publishSeen = false;
        string? publishPath = null;
        for (var i = 1; i < close; i++)
        {
            var raw = lines[i];
            if (string.IsNullOrWhiteSpace(raw) || raw.TrimStart().StartsWith("#", StringComparison.Ordinal)) continue;
            var indent = raw.Length - raw.TrimStart(' ').Length;
            var line = raw.Trim();
            if (indent == 0)
            {
                var colon = line.IndexOf(':');
                if (colon < 1) continue;
                inPublish = line[..colon] == "publish";
                if (inPublish)
                {
                    if (publishSeen) diagnostics.Add(new("YADG-PUBLISH-009", "Duplicate publish configuration.", true, $"{path}:{i + 1}"));
                    publishSeen = true;
                    if (line[(colon + 1)..].Trim().Length != 0) diagnostics.Add(new("YADG-PUBLISH-009", "publish must be a mapping to read publish.path.", true, $"{path}:{i + 1}"));
                }
                continue;
            }
            if (!inPublish || indent != 2) continue;
            var split = line.IndexOf(':');
            if (split < 1) { diagnostics.Add(new("YADG-PUBLISH-009", "Malformed publish configuration.", true, $"{path}:{i + 1}")); continue; }
            var key = line[..split].Trim();
            if (key != "path") { diagnostics.Add(new("YADG-PUBLISH-009", $"Unknown publish key '{key}'.", true, $"{path}:{i + 1}")); continue; }
            if (publishPath is not null) { diagnostics.Add(new("YADG-PUBLISH-009", "Duplicate publish.path.", true, $"{path}:{i + 1}")); continue; }
            if (!TryStringScalar(line[(split + 1)..].Trim(), out var value) || string.IsNullOrWhiteSpace(value)) diagnostics.Add(new("YADG-PUBLISH-009", "publish.path must be a non-empty string.", true, $"{path}:{i + 1}"));
            else publishPath = value;
        }
        if (publishSeen && publishPath is null && !diagnostics.Any(d => d.IsError)) diagnostics.Add(new("YADG-PUBLISH-009", "publish.path must be a non-empty string.", true, path));
        return publishPath;
    }

    public static WorkspaceValues Parse(string path, List<Diagnostic> diagnostics)
    {
        string text;
        try { text = File.ReadAllText(path, Encoding.UTF8); }
        catch (Exception ex) { diagnostics.Add(new("YADG-VALUES-001", $"Unable to read workspace marker: {ex.Message}", true, path)); return WorkspaceValues.Empty; }
        return Parse(text, path, diagnostics);
    }

    public static WorkspaceValues Parse(string text, string location, List<Diagnostic> diagnostics)
    {
        if (text.Length > 0 && text[0] == '\ufeff') text = text[1..];
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        if (lines.Length == 0 || lines[0] != "---") return WorkspaceValues.Empty;
        var close = Array.FindIndex(lines, 1, line => line == "---");
        if (close < 0)
        {
            diagnostics.Add(new("YADG-VALUES-002", "Workspace front matter has no closing delimiter.", true, location));
            return WorkspaceValues.Empty;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        var rootKeys = new HashSet<string>(StringComparer.Ordinal);
        var yadgKeys = new HashSet<string>(StringComparer.Ordinal);
        var valueKeys = new HashSet<string>(StringComparer.Ordinal);
        var producerKeys = new HashSet<string>(StringComparer.Ordinal);
        var publishKeys = new HashSet<string>(StringComparer.Ordinal);
        var mermaidKeys = new HashSet<string>(StringComparer.Ordinal);
        var markdownKeys = new HashSet<string>(StringComparer.Ordinal);
        var diagramsKeys = new HashSet<string>(StringComparer.Ordinal);
        var referenceKeys = new HashSet<string>(StringComparer.Ordinal);
        var producerArguments = new List<string>();
        string? producerExecutable = null;
        string? publishPath = null;
        var thematicBreak = "error";
        var codeInline = "error";
        var diagramsAutoDetect = false;
        var headingReferences = "number";
        var figureReferences = "number";
        var tableReferences = "number";
        var publishSeen = false;
        var section = "";
        var versionSeen = false;
        for (var i = 1; i < close; i++)
        {
            var raw = lines[i];
            if (string.IsNullOrWhiteSpace(raw)) continue;
            if (raw.Contains('\t')) { Error(diagnostics, "YADG-VALUES-003", "Tabs are not supported in workspace front matter.", location, i); continue; }
            var indent = raw.Length - raw.TrimStart(' ').Length;
            var line = raw.Trim();
            if (line.StartsWith("#", StringComparison.Ordinal)) continue;
            if (section == "presentations" && indent > 0) continue;
            if (section == "producers.mermaid.arguments" && indent == 6)
            {
                if (!line.StartsWith("-", StringComparison.Ordinal)) { Error(diagnostics, "YADG-PRODUCER-004", "Mermaid arguments must be a YAML sequence of strings.", location, i); continue; }
                var item = line[1..].Trim();
                if (!TryStringScalar(item, out var argument)) Error(diagnostics, "YADG-PRODUCER-004", "Mermaid arguments must contain only string scalars.", location, i);
                else if (argument is "-i" or "--input" or "-o" or "--output") Error(diagnostics, "YADG-PRODUCER-006", "Mermaid input/output flags are reserved by YADG.", location, i);
                else producerArguments.Add(argument);
                continue;
            }
            var colon = line.IndexOf(':');
            if (colon <= 0)
            {
                Error(diagnostics, "YADG-VALUES-005", $"Malformed front matter line '{raw}'.", location, i); continue;
            }
            var key = line[..colon].Trim();
            var rawValue = line[(colon + 1)..].Trim();
            if (rawValue.StartsWith("&", StringComparison.Ordinal) || rawValue.StartsWith("*", StringComparison.Ordinal) || rawValue.StartsWith("!", StringComparison.Ordinal) || key == "<<")
                Error(diagnostics, "YADG-VALUES-004", "Anchors, aliases, merge keys, and custom YAML tags are unsupported.", location, i);
            if (indent == 0)
            {
                if (!rootKeys.Add(key)) Error(diagnostics, "YADG-VALUES-006", $"Duplicate front matter key '{key}'.", location, i);
                if (key == "presentation") Error(diagnostics, "YADG-PRESENTATION-CONFIG-001", "The singular 'presentation' key is unsupported; use the 'presentations' sequence.", location, i);
                else if (key is not "yadg" and not "values" and not "producers" and not "publish" and not "markdown" and not "diagrams" and not "references" and not "presentations") Error(diagnostics, "YADG-VALUES-007", $"Unknown front matter key '{key}'.", location, i);
                if (rawValue.Length != 0) Error(diagnostics, "YADG-VALUES-005", $"Mapping key '{key}' must not have an inline value.", location, i);
                if (key == "publish") publishSeen = true;
                section = key; continue;
            }
            if (section == "yadg" && indent == 2)
            {
                if (!yadgKeys.Add(key)) Error(diagnostics, "YADG-VALUES-006", $"Duplicate front matter key 'yadg.{key}'.", location, i);
                if (key != "version") { Error(diagnostics, "YADG-VALUES-007", $"Unknown front matter key 'yadg.{key}'.", location, i); continue; }
                if (rawValue != "1") Error(diagnostics, "YADG-VALUES-008", "Workspace front matter version must be 1.", location, i);
                versionSeen = true; continue;
            }
            if (section == "values" && indent == 2)
            {
                if (!Id.IsMatch(key)) { Error(diagnostics, "YADG-VALUES-009", $"Invalid workspace value ID '{key}'.", location, i); continue; }
                if (!valueKeys.Add(key)) { Error(diagnostics, "YADG-VALUES-006", $"Duplicate workspace value '{key}'.", location, i); continue; }
                if (rawValue is "|" or ">" || rawValue.StartsWith("|", StringComparison.Ordinal) || rawValue.StartsWith(">", StringComparison.Ordinal))
                { Error(diagnostics, "YADG-VALUES-010", $"Workspace value '{key}' must be a single-line string.", location, i); continue; }
                if (rawValue.Length == 0 || rawValue.StartsWith("[", StringComparison.Ordinal) || rawValue.StartsWith("{", StringComparison.Ordinal))
                { Error(diagnostics, "YADG-VALUES-011", $"Workspace value '{key}' must be a YAML string scalar.", location, i); continue; }
                if (!TryScalar(rawValue, out var value) || value.Contains('\n') || value.Contains('\r'))
                { Error(diagnostics, "YADG-VALUES-011", $"Workspace value '{key}' must be a single-line YAML string scalar.", location, i); continue; }
                values[key] = value; continue;
            }
            if (section == "markdown" && indent == 2)
            {
                if (!markdownKeys.Add(key)) { Error(diagnostics, "YADG-VALUES-006", $"Duplicate markdown key '{key}'.", location, i); continue; }
                if (!TryStringScalar(rawValue, out var mode)) { Error(diagnostics, "YADG-VALUES-007", $"markdown.{key} must be a supported string mode.", location, i); continue; }
                if (key == "thematicBreak" && mode is "error" or "ignore") thematicBreak = mode;
                else if (key == "codeInline" && mode is "error" or "ignore" or "style") codeInline = mode;
                else Error(diagnostics, "YADG-VALUES-007", key is "thematicBreak" or "codeInline" ? $"Unsupported markdown.{key} mode '{mode}'." : $"Unknown markdown key '{key}'.", location, i);
                continue;
            }
            if (section == "diagrams" && indent == 2)
            {
                if (!diagramsKeys.Add(key)) { Error(diagnostics, "YADG-VALUES-006", $"Duplicate diagrams key '{key}'.", location, i); continue; }
                if (key != "autoDetect") { Error(diagnostics, "YADG-VALUES-007", $"Unknown diagrams key '{key}'.", location, i); continue; }
                if (rawValue.Equals("true", StringComparison.OrdinalIgnoreCase)) diagramsAutoDetect = true;
                else if (rawValue.Equals("false", StringComparison.OrdinalIgnoreCase)) diagramsAutoDetect = false;
                else Error(diagnostics, "YADG-VALUES-007", "diagrams.autoDetect must be a YAML boolean.", location, i);
                continue;
            }
            if (section == "references" && indent == 2)
            {
                if (!referenceKeys.Add(key)) { Error(diagnostics, "YADG-VALUES-006", $"Duplicate references key '{key}'.", location, i); continue; }
                if (key is not "headings" and not "figures" and not "tables") { Error(diagnostics, "YADG-VALUES-007", $"Unknown references key '{key}'.", location, i); continue; }
                if (!TryStringScalar(rawValue, out var mode)) { Error(diagnostics, "YADG-VALUES-007", $"references.{key} must be a supported string mode.", location, i); continue; }
                var valid = key == "headings" ? mode == "number" : mode is "number" or "label-number";
                if (!valid) { Error(diagnostics, "YADG-VALUES-007", key == "headings" ? "references.headings supports only 'number'." : $"references.{key} supports only 'number' or 'label-number'.", location, i); continue; }
                if (key == "headings") headingReferences = mode;
                else if (key == "figures") figureReferences = mode;
                else tableReferences = mode;
                continue;
            }
            if (section == "producers" && indent == 2)
            {
                if (key != "mermaid") { Error(diagnostics, "YADG-PRODUCER-001", $"Unsupported producer kind '{key}'.", location, i); continue; }
                if (!producerKeys.Add(key)) Error(diagnostics, "YADG-VALUES-006", "Duplicate producer configuration 'mermaid'.", location, i);
                if (rawValue.Length != 0) Error(diagnostics, "YADG-PRODUCER-002", "Producer kind must be a mapping.", location, i);
                section = "producers.mermaid"; continue;
            }
            if (section == "publish" && indent == 2)
            {
                if (!publishKeys.Add(key)) { Error(diagnostics, "YADG-VALUES-006", $"Duplicate publish key '{key}'.", location, i); continue; }
                if (key != "path") { Error(diagnostics, "YADG-VALUES-007", $"Unknown publish key '{key}'.", location, i); continue; }
                if (!TryStringScalar(rawValue, out var path) || string.IsNullOrWhiteSpace(path)) Error(diagnostics, "YADG-PUBLISH-001", "publish.path must be a non-empty string.", location, i);
                else publishPath = path;
                continue;
            }
            if (section == "producers.mermaid" && indent == 4)
            {
                if (!mermaidKeys.Add(key)) { Error(diagnostics, "YADG-VALUES-006", $"Duplicate producer key 'mermaid.{key}'.", location, i); continue; }
                if (key == "executable")
                {
                    if (!TryStringScalar(rawValue, out var executable) || string.IsNullOrWhiteSpace(executable)) Error(diagnostics, "YADG-PRODUCER-003", "Mermaid executable must be a non-empty string.", location, i);
                    else producerExecutable = executable;
                    continue;
                }
                if (key == "arguments")
                {
                    if (rawValue == "[]") { section = "producers.mermaid"; continue; }
                    if (rawValue.Length != 0) Error(diagnostics, "YADG-PRODUCER-004", "Mermaid arguments must be a YAML sequence of strings.", location, i);
                    section = "producers.mermaid.arguments"; continue;
                }
                Error(diagnostics, "YADG-PRODUCER-005", $"Unknown Mermaid producer key '{key}'.", location, i); continue;
            }
            Error(diagnostics, "YADG-VALUES-007", $"Unsupported or incorrectly indented front matter key '{key}'.", location, i);
        }
        if (!rootKeys.Contains("yadg") || !versionSeen) diagnostics.Add(new("YADG-VALUES-008", "Workspace front matter must declare yadg.version: 1.", true, location));
        if (producerKeys.Contains("mermaid") && producerExecutable is null) diagnostics.Add(new("YADG-PRODUCER-003", "Mermaid producer configuration requires executable.", true, location));
        if (publishSeen && publishPath is null) diagnostics.Add(new("YADG-PUBLISH-001", "publish.path must be a non-empty string.", true, location));
        var presentations = ParsePresentations(lines, close, location, diagnostics);
        return new(values) { PublishPath = publishPath, ThematicBreakPolicy = thematicBreak, CodeInlinePolicy = codeInline, DiagramsAutoDetect = diagramsAutoDetect, References = new(headingReferences, figureReferences, tableReferences), Producers = producerExecutable is null ? new Dictionary<string, MermaidProducerConfiguration>(StringComparer.Ordinal) : new Dictionary<string, MermaidProducerConfiguration>(StringComparer.Ordinal) { ["mermaid"] = new(producerExecutable, producerArguments) }, Presentations = presentations };
    }

    private static IReadOnlyList<PresentationDeclaration> ParsePresentations(string[] lines, int close, string location, List<Diagnostic> diagnostics)
    {
        var start = -1;
        for (var i = 1; i < close; i++)
        {
            var line = lines[i];
            if (line.Length == 0 || char.IsWhiteSpace(line[0])) continue;
            if (line.StartsWith("presentations:", StringComparison.Ordinal)) { start = i; break; }
        }
        if (start < 0) return Array.Empty<PresentationDeclaration>();
        if (lines[start]["presentations:".Length..].Trim().Length != 0)
        {
            Error(diagnostics, "YADG-PRESENTATION-CONFIG-002", "presentations must be a non-empty sequence.", location, start);
            return Array.Empty<PresentationDeclaration>();
        }

        var results = new List<PresentationDeclaration>();
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        var seenOutputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, string> fields = new(StringComparer.Ordinal);
        List<string>? sources = null;
        var entryLine = start;
        var inSources = false;

        void FinishEntry()
        {
            if (fields.Count == 0 && sources is null) return;
            var allowed = new HashSet<string>(new[] { "id", "template", "sources", "output" }, StringComparer.Ordinal);
            foreach (var key in fields.Keys.Where(key => !allowed.Contains(key)))
                Error(diagnostics, "YADG-PRESENTATION-CONFIG-003", $"Unknown presentation key '{key}'.", location, entryLine);
            foreach (var key in allowed.Where(key => key != "sources" && !fields.ContainsKey(key)))
                Error(diagnostics, "YADG-PRESENTATION-CONFIG-004", $"Presentation entry requires '{key}'.", location, entryLine);
            if (sources is null || sources.Count == 0)
                Error(diagnostics, "YADG-PRESENTATION-CONFIG-004", "Presentation 'sources' must be a non-empty ordered sequence of strings.", location, entryLine);

            if (fields.TryGetValue("id", out var id) && !Id.IsMatch(id))
                Error(diagnostics, "YADG-PRESENTATION-CONFIG-005", "Presentation id must match [A-Za-z][A-Za-z0-9_-]*.", location, entryLine);
            if (fields.TryGetValue("id", out id) && Id.IsMatch(id) && !seenIds.Add(id))
                Error(diagnostics, "YADG-PRESENTATION-CONFIG-006", $"Duplicate presentation id '{id}'.", location, entryLine);
            if (fields.TryGetValue("output", out var output))
            {
                var stem = Path.GetFileNameWithoutExtension(output);
                var reservedStem = stem.Split('.')[0].ToUpperInvariant();
                var reservedDevice = reservedStem is "CON" or "PRN" or "AUX" or "NUL" || Regex.IsMatch(reservedStem, @"^(COM|LPT)[1-9]$");
                if (string.IsNullOrWhiteSpace(output) || Path.GetFileName(output) != output || !output.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) || output is "." or ".." || stem.Length == 0 || reservedDevice || output.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                    Error(diagnostics, "YADG-PRESENTATION-CONFIG-007", "Presentation output must be a valid .pptx basename without path components.", location, entryLine);
                else if (!seenOutputs.Add(output)) Error(diagnostics, "YADG-PRESENTATION-CONFIG-008", $"Duplicate presentation output basename '{output}' (case-insensitive).", location, entryLine);
            }
            if (fields.TryGetValue("id", out id) && fields.TryGetValue("template", out var template) && fields.TryGetValue("output", out output) && sources is { Count: > 0 })
                results.Add(new(id, template, sources.ToArray(), output));
            fields = new(StringComparer.Ordinal);
            sources = null;
            inSources = false;
        }

        for (var i = start + 1; i < close; i++)
        {
            var raw = lines[i];
            if (string.IsNullOrWhiteSpace(raw) || raw.TrimStart().StartsWith("#", StringComparison.Ordinal)) continue;
            if (raw.Contains('\t')) { Error(diagnostics, "YADG-VALUES-003", "Tabs are not supported in workspace front matter.", location, i); continue; }
            var indent = raw.Length - raw.TrimStart(' ').Length;
            var line = raw.Trim();
            if (indent == 0) break;
            if (indent == 2 && line.StartsWith("- ", StringComparison.Ordinal))
            {
                FinishEntry();
                entryLine = i;
                inSources = false;
                ParseField(line[2..], i, allowSourceKey: true);
                continue;
            }
            if (indent == 4)
            {
                ParseField(line, i, allowSourceKey: true);
                continue;
            }
            if (indent == 6 && inSources && line.StartsWith("- ", StringComparison.Ordinal))
            {
                var scalar = line[2..].Trim();
                if (!TryStrictStringScalar(scalar, out var source) || string.IsNullOrWhiteSpace(source)) Error(diagnostics, "YADG-PRESENTATION-CONFIG-009", "Presentation sources must be non-empty YAML string scalars.", location, i);
                else sources!.Add(source);
                continue;
            }
            Error(diagnostics, "YADG-PRESENTATION-CONFIG-010", "Malformed presentations sequence; use two-space deck entries and a six-space sources sequence.", location, i);
        }
        FinishEntry();
        if (results.Count == 0 && !diagnostics.Any(d => d.IsError && (d.Code.StartsWith("YADG-PRESENTATION-CONFIG-", StringComparison.Ordinal))))
            Error(diagnostics, "YADG-PRESENTATION-CONFIG-002", "presentations must contain at least one deck declaration.", location, start);
        return results;

        void ParseField(string text, int lineNumber, bool allowSourceKey)
        {
            var colon = text.IndexOf(':');
            if (colon <= 0) { Error(diagnostics, "YADG-PRESENTATION-CONFIG-010", "Malformed presentation mapping entry.", location, lineNumber); return; }
            var key = text[..colon].Trim();
            var rawValue = text[(colon + 1)..].Trim();
            if (key == "sources")
            {
                if (!allowSourceKey || rawValue.Length != 0 || sources is not null) { Error(diagnostics, "YADG-PRESENTATION-CONFIG-010", "sources must be one ordered YAML sequence.", location, lineNumber); return; }
                sources = new List<string>();
                inSources = true;
                return;
            }
            if (fields.ContainsKey(key)) { Error(diagnostics, "YADG-PRESENTATION-CONFIG-011", $"Duplicate presentation key '{key}'.", location, lineNumber); return; }
            if (!TryStrictStringScalar(rawValue, out var value) || string.IsNullOrWhiteSpace(value)) { Error(diagnostics, "YADG-PRESENTATION-CONFIG-009", $"Presentation '{key}' must be a non-empty YAML string scalar.", location, lineNumber); return; }
            fields[key] = value;
            inSources = false;
        }
    }

    private static bool TryStrictStringScalar(string raw, out string value)
    {
        value = string.Empty;
        if (!TryScalar(raw, out value)) return false;
        if (raw.StartsWith('"') && raw.EndsWith('"') && raw.Length >= 2) return true;
        if (raw.StartsWith('\'') && raw.EndsWith('\'') && raw.Length >= 2) return true;
        return !string.IsNullOrWhiteSpace(value);
    }

    private static bool TryScalar(string raw, out string value)
    {
        value = "";
        if (raw.StartsWith('"') && raw.EndsWith('"') && raw.Length >= 2)
        {
            try { value = Regex.Unescape(raw[1..^1]); return true; } catch { return false; }
        }
        if (raw.StartsWith('\'') && raw.EndsWith('\'') && raw.Length >= 2) { value = raw[1..^1].Replace("''", "'"); return true; }
        if (raw is "null" or "Null" or "NULL" or "true" or "True" or "TRUE" or "false" or "False" or "FALSE" || Number.IsMatch(raw) || DateLike.IsMatch(raw)) return false;
        if (raw.StartsWith("- ", StringComparison.Ordinal) || raw.Contains(" #", StringComparison.Ordinal)) return false;
        value = raw; return true;
    }

    private static bool TryStringScalar(string raw, out string value)
    {
        value = string.Empty;
        if (raw.Length == 0 || raw is "null" or "Null" or "NULL" or "true" or "false") return false;
        if (raw.StartsWith('"') && raw.EndsWith('"') && raw.Length >= 2) { try { value = Regex.Unescape(raw[1..^1]); return true; } catch { return false; } }
        if (raw.StartsWith('\'') && raw.EndsWith('\'') && raw.Length >= 2) { value = raw[1..^1].Replace("''", "'"); return true; }
        if (raw.StartsWith("[", StringComparison.Ordinal) || raw.StartsWith("{", StringComparison.Ordinal) || raw.Contains(" #", StringComparison.Ordinal)) return false;
        value = raw; return true;
    }

    private static void Error(List<Diagnostic> diagnostics, string code, string message, string location, int line)
    {
        var column = 1;
        try
        {
            if (File.Exists(location))
            {
                var sourceLine = File.ReadLines(location).Skip(line).FirstOrDefault();
                if (sourceLine is not null) column = sourceLine.TakeWhile(char.IsWhiteSpace).Count() + 1;
            }
        }
        catch { }
        diagnostics.Add(new(code, message, true, $"{location}:{line + 1}:{column}"));
    }
}
