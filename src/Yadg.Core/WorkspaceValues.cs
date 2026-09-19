using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Yadg.Core;

public sealed record WorkspaceValues(IReadOnlyDictionary<string, string> Values)
{
    public static WorkspaceValues Empty { get; } = new(new Dictionary<string, string>(StringComparer.Ordinal));
}

public static class WorkspaceValuesParser
{
    private static readonly Regex Id = new("^[A-Za-z][A-Za-z0-9_-]*$", RegexOptions.Compiled);
    private static readonly Regex Number = new("^[+-]?(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?$", RegexOptions.Compiled);
    private static readonly Regex DateLike = new("^[0-9]{4}-[0-9]{2}-[0-9]{2}(?:[Tt ].*)?$", RegexOptions.Compiled);

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
                if (key is not "yadg" and not "values") Error(diagnostics, "YADG-VALUES-007", $"Unknown front matter key '{key}'.", location, i);
                if (rawValue.Length != 0) Error(diagnostics, "YADG-VALUES-005", $"Mapping key '{key}' must not have an inline value.", location, i);
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
            Error(diagnostics, "YADG-VALUES-007", $"Unsupported or incorrectly indented front matter key '{key}'.", location, i);
        }
        if (!rootKeys.Contains("yadg") || !versionSeen) diagnostics.Add(new("YADG-VALUES-008", "Workspace front matter must declare yadg.version: 1.", true, location));
        return new(values);
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

    private static void Error(List<Diagnostic> diagnostics, string code, string message, string location, int line) => diagnostics.Add(new(code, message, true, $"{location}:{line + 1}"));
}
