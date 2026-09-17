using System.Collections.ObjectModel;

namespace Yadg.Core;

public sealed record Diagnostic(string Code, string Message, bool IsError = true)
{
    public override string ToString() => $"{Code}: {Message}";
}

public sealed record YadgDocument(IReadOnlyList<YadgSection> Sections)
{
    public YadgSection? FindSection(string id) => Sections.SingleOrDefault(s => s.Id == id);
}

public sealed record YadgSection(string Id, string Heading, int Level, IReadOnlyList<string> BodyLines)
{
    public string SectionText => string.Join(Environment.NewLine, new[] { Heading }.Concat(BodyLines));
    public string ContentText => string.Join(Environment.NewLine, BodyLines);
}

public enum SectionSelection { Section, Content }

public sealed record SectionReference(string Id, SectionSelection Selection)
{
    public static bool TryParseTag(string value, out SectionReference? reference, out Diagnostic? diagnostic)
    {
        reference = null;
        diagnostic = null;
        const string prefix = "{{yadg:section:content:";
        if (!value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) || !value.EndsWith("}}", StringComparison.Ordinal))
        {
            diagnostic = new("YADG-TAG-001", $"Malformed tag '{value}'. Expected {{yadg:section:content:<stable-id>}}.");
            return false;
        }
        var id = value[prefix.Length..^2];
        if (string.IsNullOrWhiteSpace(id) || id.Any(char.IsWhiteSpace))
        {
            diagnostic = new("YADG-TAG-001", $"Malformed tag '{value}': stable ID is empty or contains whitespace.");
            return false;
        }
        reference = new(id, SectionSelection.Content);
        return true;
    }
}

public sealed record ParseResult(YadgDocument? Document, IReadOnlyList<Diagnostic> Diagnostics)
{
    public bool IsValid => Diagnostics.All(d => !d.IsError) && Document is not null;
}
