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
        var command = new Command("check", "Parse Markdown and validate visible template references.");
        var markdown = new Option<FileInfo>("--markdown") { IsRequired = true };
        var template = new Option<FileInfo>("--template") { IsRequired = true };
        command.AddOption(markdown); command.AddOption(template);
        command.SetHandler((FileInfo md, FileInfo docx) =>
        {
            var result = MarkdownDocumentParser.Parse(File.ReadAllText(md.FullName));
            foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine(diagnostic);
            if (!result.IsValid) { fail(); return; }
            try
            {
                var tags = WordAuthoring.FindTags(docx.FullName);
                if (tags.Count == 0) throw new WordAuthoringException(new("YADG-TAG-002", "Template contains no supported visible YADG section-content tag."));
                foreach (var tag in tags)
                {
                    if (!SectionReference.TryParseTag(tag, out var reference, out var tagDiagnostic)) throw new WordAuthoringException(tagDiagnostic!);
                    var resolved = MarkdownDocumentParser.Resolve(result.Document!, reference!);
                    if (resolved.Diagnostic is not null) throw new WordAuthoringException(resolved.Diagnostic);
                }
                Console.WriteLine($"check: valid ({tags.Count} visible tag(s), {result.Document!.Sections.Count} section(s))");
            }
            catch (WordAuthoringException ex) { Console.Error.WriteLine(ex.Diagnostic); fail(); }
        }, markdown, template);
        return command;
    }

    private static Command CreateBuildCommand(Action fail)
    {
        var command = new Command("build", "Author a DOCX from Markdown and a prepared template without Microsoft Word.");
        var markdown = new Option<FileInfo>("--markdown") { IsRequired = true };
        var template = new Option<FileInfo>("--template") { IsRequired = true };
        var output = new Option<FileInfo>("--output") { IsRequired = true };
        command.AddOption(markdown); command.AddOption(template); command.AddOption(output);
        command.SetHandler((FileInfo md, FileInfo docx, FileInfo resultPath) =>
        {
            var parsed = MarkdownDocumentParser.Parse(File.ReadAllText(md.FullName));
            foreach (var diagnostic in parsed.Diagnostics) Console.Error.WriteLine(diagnostic);
            if (!parsed.IsValid) { fail(); return; }
            try
            {
                WordAuthoring.Author(docx.FullName, resultPath.FullName, parsed.Document!);
                Console.WriteLine($"build: wrote {resultPath.FullName}");
            }
            catch (WordAuthoringException ex) { Console.Error.WriteLine(ex.Diagnostic); fail(); }
        }, markdown, template, output);
        return command;
    }
}
