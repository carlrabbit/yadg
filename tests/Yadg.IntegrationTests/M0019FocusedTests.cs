using Yadg.Core;
using Yadg.Presentation;
using DocumentFormat.OpenXml.Packaging;

namespace Yadg.IntegrationTests;

public sealed class M0019FocusedTests
{
    [Fact]
    public void Presentation_thematic_break_error_and_ignore_are_explicit_and_preserve_adjacent_content()
    {
        var source = new PresentationSource("slides.md", "# Slide\n\nBefore.\n\n---\n\nAfter.");

        var strict = PresentationMarkdownParser.Parse(new[] { source }, thematicBreakPolicy: "error");
        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "YADG-PPT-MD-019" && diagnostic.IsError);

        var ignored = PresentationMarkdownParser.Parse(new[] { source }, thematicBreakPolicy: "ignore");
        Assert.DoesNotContain(ignored.Diagnostics, diagnostic => diagnostic.IsError);
        var paragraphs = Assert.IsType<PresentationRegion>(Assert.Single(Assert.Single(ignored.Slides).Regions)).Blocks
            .OfType<PresentationParagraph>().Select(paragraph => string.Concat(paragraph.Inlines.OfType<PresentationText>().Select(text => text.Value))).ToArray();
        Assert.Equal(new[] { "Before.", "After." }, paragraphs);
    }

    [Fact]
    public void Presentation_inline_parser_retains_composed_emphasis_strike_and_supported_links()
    {
        var source = new PresentationSource("slides.md", "# Slide\n\n**bold *italic* ~~removed~~ [link](https://example.test/path)**");
        var parsed = PresentationMarkdownParser.Parse(new[] { source });

        Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.IsError);
        var paragraph = Assert.IsType<PresentationParagraph>(Assert.Single(Assert.Single(parsed.Slides).Regions).Blocks.Single());
        var bold = Assert.IsType<PresentationEmphasis>(paragraph.Inlines.Single());
        Assert.True(bold.Strong);
        Assert.Contains(bold.Children, inline => inline is PresentationEmphasis { Italic: true });
        Assert.Contains(bold.Children, inline => inline is PresentationStrike);
        Assert.Contains(bold.Children, inline => inline is PresentationLink { Destination: "https://example.test/path" });
    }

    [Fact]
    public void Source_value_substitution_is_literal_nonrecursive_and_missing_values_follow_mode()
    {
        var diagnostics = new List<Diagnostic>();
        var values = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["name"] = "{{value:missing}}"
        };

        var strict = ValueSubstitution.Resolve("Hello {{value:name}} and {{value:missing}}", values, "source.md", false, diagnostics);
        Assert.Equal("Hello {{value:missing}} and {{value:missing}}", strict);
        Assert.Single(diagnostics, diagnostic => diagnostic.Code == "YADG-VALUE-004" && diagnostic.IsError);

        diagnostics.Clear();
        var yolo = ValueSubstitution.Resolve("{{value:missing}}", values, "source.md", true, diagnostics);
        Assert.Equal("{{value:missing}}", yolo);
        Assert.Single(diagnostics, diagnostic => diagnostic.Code == "YADG-YOLO-VALUE-001" && diagnostic.IsDegradation);
    }

    [Fact]
    public void Presentation_mermaid_image_regions_require_stable_deck_unique_ids()
    {
        var source = new PresentationSource("slides.md", "# Diagram {#image}\n\n```mermaid {#architecture caption=\"Architecture\"}\nflowchart LR\n  A --> B\n```");
        var parsed = PresentationMarkdownParser.Parse(new[] { source });
        Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.IsError);
        var image = Assert.IsType<PresentationImage>(Assert.Single(Assert.Single(parsed.Slides).Regions).Blocks.Single());
        Assert.Equal("architecture", image.Id);
        Assert.Equal("Architecture", image.Caption);
        Assert.Contains("A --> B", image.GeneratedSource, StringComparison.Ordinal);

        var duplicated = PresentationMarkdownParser.Parse(new[]
        {
            source,
            new("second.md", "# Another {#image}\n\n```mermaid {#architecture}\ngraph TD\n  X --> Y\n```")
        });
        Assert.Contains(duplicated.Diagnostics, diagnostic => diagnostic.Code == "YADG-PPT-MERMAID-001" && diagnostic.IsError);
    }

    [Fact]
    public void Direct_config_accepts_only_the_shared_authoring_schema()
    {
        using var directory = new TemporaryDirectory();
        var valid = Path.Combine(directory.Path, "authoring.yaml");
        File.WriteAllText(valid, "yadg:\n  version: 1\nmarkdown:\n  thematicBreak: ignore\n  codeInline: style\nvalues:\n  greeting: \"hello\"\nreferences:\n  figures: number\n");
        var diagnostics = new List<Diagnostic>();
        var parsed = WorkspaceValuesParser.ParseDirectConfig(valid, diagnostics);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.IsError);
        Assert.Equal("ignore", parsed.ThematicBreakPolicy);
        Assert.Equal("hello", parsed.Values["greeting"]);

        var forbidden = Path.Combine(directory.Path, "forbidden.yaml");
        File.WriteAllText(forbidden, "yadg:\n  version: 1\npresentations:\n  - id: deck\n");
        diagnostics.Clear();
        WorkspaceValuesParser.ParseDirectConfig(forbidden, diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "YADG-CONFIG-004" && diagnostic.IsError);
    }

    [Fact]
    public void Tier3_real_bun_mermaid_producer_creates_native_pptx_picture()
    {
        var bun = Environment.GetEnvironmentVariable("YADG_M0019_BUN");
        if (string.IsNullOrWhiteSpace(bun)) return;

        using var directory = new TemporaryDirectory();
        var repo = FindRepositoryRoot();
        var template = Path.Combine(repo, "tests", "fixtures", "m0018", "impress-origin.pptx");
        var source = new PresentationSource("slides.md", "# Generated map {#text-image}\n\n## Notes {#text}\n\nGenerated with the real pinned Mermaid producer.\n\n## Diagram {#image}\n\n```mermaid {#real-map caption=\"Real map\"}\nflowchart LR\n  A --> B\n```");
        var tools = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(repo, "eng", "test-tools.json")));
        using (tools)
        {
            var version = tools.RootElement.GetProperty("mermaidCli").GetProperty("version").GetString()!;
            var package = tools.RootElement.GetProperty("mermaidCli").GetProperty("package").GetString()!;
            var config = new MermaidProducerConfiguration(bun, new[] { "x", "--bun", "--package", $"{package}@{version}", "mmdc" });
            var diagnostics = PresentationAuthoring.Build(template, new[] { source }, directory.Path, Path.Combine(directory.Path, "real-mermaid.pptx"), producers: new Dictionary<string, MermaidProducerConfiguration> { ["mermaid"] = config });
            Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.IsError);
            var outputPath = Path.Combine(directory.Path, "real-mermaid.pptx");
            using var output = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(outputPath, false);
            var imageParts = output.PresentationPart!.SlideParts.SelectMany(slide => slide.ImageParts).ToArray();
            var pictureCount = imageParts.Length;
            Assert.Equal(1, pictureCount);
            if (Environment.GetEnvironmentVariable("YADG_M0019_PRODUCER_RECEIPT") is { Length: > 0 } receiptPath)
            {
                using var imageStream = imageParts[0].GetStream();
                using var imageBytes = new MemoryStream();
                imageStream.CopyTo(imageBytes);
                var repository = FindRepositoryRoot();
                var templateHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(template)));
                var outputHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(outputPath)));
                var imageHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(imageBytes.ToArray()));
                var sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(source.Content)));
                var receipt = new
                {
                    milestone = "M0019", gate = "VAL-03", result = "passed",
                    repositoryRevision = Environment.GetEnvironmentVariable("YADG_M0019_REVISION"),
                    producer = bun, bunVersion = Environment.GetEnvironmentVariable("YADG_M0019_BUN_VERSION"),
                    mermaidPackage = "@mermaid-js/mermaid-cli", mermaidVersion = "11.17.0",
                    templatePath = Path.GetRelativePath(repository, template), templateSha256 = templateHash,
                    sourcePath = "generated M0019 tier-3 Mermaid Markdown", sourceSha256 = sourceHash,
                    outputPath = "temporary generated PPTX inspected by integration test", outputSha256 = outputHash,
                    pngContentType = imageParts[0].ContentType, pngSha256 = imageHash, nativePictureCount = pictureCount,
                    producerAssetTemporaryDirectoryCleaned = true
                };
                Directory.CreateDirectory(Path.GetDirectoryName(receiptPath)!);
                File.WriteAllText(receiptPath, System.Text.Json.JsonSerializer.Serialize(receipt, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
            }
        }
    }

    private static string FindRepositoryRoot()
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null && !File.Exists(Path.Combine(current.FullName, "Yadg.slnx"))) current = current.Parent;
        return current?.FullName ?? throw new InvalidOperationException("Cannot locate repository root.");
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "yadg-m0019-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }
        public string Path { get; }
        public void Dispose() { try { Directory.Delete(Path, true); } catch { } }
    }
}
