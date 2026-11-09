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

        var recovered = PresentationMarkdownParser.Parse(new[] { source }, thematicBreakPolicy: "error", yolo: true);
        Assert.DoesNotContain(recovered.Diagnostics, diagnostic => diagnostic.IsError);
        Assert.Contains(recovered.Diagnostics, diagnostic => diagnostic.Code == "YADG-YOLO-MD-001" && diagnostic.IsDegradation);
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
    public void Presentation_inline_code_obeys_error_ignore_and_style_policies()
    {
        var source = new PresentationSource("slides.md", "# Slide\n\nA `literal` span.");
        var strict = PresentationMarkdownParser.Parse(new[] { source }, inlineCodePolicy: "error");
        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "YADG-PPT-MD-020" && diagnostic.IsError);

        var ignored = PresentationMarkdownParser.Parse(new[] { source }, inlineCodePolicy: "ignore");
        Assert.DoesNotContain(ignored.Diagnostics, diagnostic => diagnostic.IsError);
        var ignoredParagraph = Assert.IsType<PresentationParagraph>(Assert.Single(Assert.Single(ignored.Slides).Regions).Blocks.Single());
        Assert.Contains(ignoredParagraph.Inlines, inline => inline is PresentationText { Value: "literal" });

        var styled = PresentationMarkdownParser.Parse(new[] { source }, inlineCodePolicy: "style");
        var styledParagraph = Assert.IsType<PresentationParagraph>(Assert.Single(Assert.Single(styled.Slides).Regions).Blocks.Single());
        Assert.Contains(styledParagraph.Inlines, inline => inline is PresentationCode { Value: "literal" });
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

        diagnostics.Clear();
        ValueSubstitution.Resolve("{{value:bad id}}", values, "source.md", true, diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "YADG-VALUE-003" && diagnostic.IsError);
    }

    [Fact]
    public void Source_values_substitute_figure_captions_without_changing_figure_or_reference_identity()
    {
        const string mermaidSource = "flowchart LR\n  A[\"{{value:caption}}\"] --> B";
        var markdown = "# Figures {#figures}\n\n![System {{value:caption}}](images/system.png){#system-figure}\n\nFigure [@system-figure].\n\n| Name | Value |\n|---|---|\n| A | B |\n{#system-table caption=\"Table {{value:caption}}\"}\n\n```mermaid {#generated-figure caption=\"Generated {{value:caption}}\"}\n" + mermaidSource + "\n```";
        var parsed = MarkdownDocumentParser.Parse(markdown, "source.md");
        Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.IsError);
        var diagnostics = new List<Diagnostic>();
        var values = new Dictionary<string, string>(StringComparer.Ordinal) { ["caption"] = "Q4 {{value:later}}" };
        var document = ValueSubstitution.ApplyToDocument(parsed.Document!, values, false, diagnostics, "source.md");

        var fileFigure = Assert.IsType<YadgFigure>(document.FindFigure("system-figure"));
        Assert.Equal("System Q4 {{value:later}}", fileFigure.AltText);
        Assert.Equal("system-figure", fileFigure.Id);
        Assert.Equal("images/system.png", fileFigure.AssetPath);
        Assert.Same(fileFigure, Assert.IsType<YadgFigure>(document.Blocks.Single(block => block is YadgFigure { Id: "system-figure" })));
        Assert.Equal("Table Q4 {{value:later}}", Assert.IsType<YadgTable>(document.FindTable("system-table")).Caption);
        Assert.Contains(document.Blocks.OfType<YadgParagraph>().SelectMany(paragraph => paragraph.Inlines), inline => inline is YadgReference { Id: "system-figure" });

        var generatedFigure = Assert.IsType<YadgFigure>(document.FindFigure("generated-figure"));
        Assert.Equal("Generated Q4 {{value:later}}", generatedFigure.AltText);
        Assert.Equal(mermaidSource, generatedFigure.GeneratedSource);
        Assert.Empty(diagnostics);
    }

    [Fact]
    public void Source_figure_caption_missing_values_are_fatal_or_visible_yolo_degradations()
    {
        var parsed = MarkdownDocumentParser.Parse("![Missing {{value:unknown}}](figure.png){#figure-id}", "figures.md");
        Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.IsError);

        var strictDiagnostics = new List<Diagnostic>();
        var strict = ValueSubstitution.ApplyToDocument(parsed.Document!, new Dictionary<string, string>(), false, strictDiagnostics, "figures.md");
        Assert.Equal("Missing {{value:unknown}}", Assert.IsType<YadgFigure>(strict.FindFigure("figure-id")).AltText);
        Assert.Single(strictDiagnostics, diagnostic => diagnostic.Code == "YADG-VALUE-004" && diagnostic.IsError && diagnostic.Location!.Contains("figure caption", StringComparison.Ordinal));

        var yoloDiagnostics = new List<Diagnostic>();
        var yolo = ValueSubstitution.ApplyToDocument(parsed.Document!, new Dictionary<string, string>(), true, yoloDiagnostics, "figures.md");
        Assert.Equal("Missing {{value:unknown}}", Assert.IsType<YadgFigure>(yolo.FindFigure("figure-id")).AltText);
        Assert.Single(yoloDiagnostics, diagnostic => diagnostic.Code == "YADG-YOLO-VALUE-001" && diagnostic.IsDegradation && !diagnostic.IsError);
    }

    [Fact]
    public void Mermaid_failed_producer_paths_clean_owned_directories_including_timeout_and_exceptions()
    {
        using var directory = new TemporaryDirectory();
        var exitScript = Path.Combine(directory.Path, "exit.ps1");
        var invalidScript = Path.Combine(directory.Path, "invalid.ps1");
        var missingScript = Path.Combine(directory.Path, "missing.ps1");
        var sleepScript = Path.Combine(directory.Path, "sleep.ps1");
        File.WriteAllText(exitScript, "exit 7");
        File.WriteAllText(invalidScript, "[IO.File]::WriteAllText($args[3], 'not a png')");
        File.WriteAllText(missingScript, "exit 0");
        File.WriteAllText(sleepScript, "Start-Sleep -Seconds 90");

        AssertFailedProducerIsCleaned("YADG-PRODUCER-010", new("yadg-missing-producer-" + Guid.NewGuid().ToString("N"), Array.Empty<string>()), directory.Path);
        AssertFailedProducerIsCleaned("YADG-PRODUCER-012", new("powershell.exe", new[] { "-NoProfile", "-File", exitScript }), directory.Path);
        AssertFailedProducerIsCleaned("YADG-PRODUCER-013", new("powershell.exe", new[] { "-NoProfile", "-File", invalidScript }), directory.Path);
        AssertFailedProducerIsCleaned("YADG-PRODUCER-013", new("powershell.exe", new[] { "-NoProfile", "-File", missingScript }), directory.Path);
        AssertFailedProducerIsCleaned("YADG-PRODUCER-010", new("powershell.exe", new[] { "-NoProfile", "-File", missingScript }), Path.Combine(directory.Path, "missing-working-directory"));
        AssertFailedProducerIsCleaned("YADG-PRODUCER-011", new("powershell.exe", new[] { "-NoProfile", "-File", sleepScript }), directory.Path, 100);
    }

    [Fact]
    public void Mermaid_producer_failure_preserves_strict_output_and_yolo_build_keeps_visible_placeholder()
    {
        using var directory = new TemporaryDirectory();
        var template = Path.Combine(FindRepositoryRoot(), "tests", "fixtures", "m0018", "impress-origin.pptx");
        var source = new PresentationSource(Path.Combine(directory.Path, "slides.md"), "# Failed diagram {#text-image}\n\n## Notes {#text}\n\nProducer recovery test.\n\n## Diagram {#image}\n\n```mermaid {#producer-failure caption=\"Failure\"}\nflowchart LR\n  A --> B\n```");
        var producers = new Dictionary<string, MermaidProducerConfiguration>(StringComparer.Ordinal)
        {
            ["mermaid"] = new("cmd.exe", new[] { "/d", "/c", "exit 7" })
        };
        var output = Path.Combine(directory.Path, "committed.pptx");
        File.WriteAllText(output, "committed sentinel");

        var strict = PresentationAuthoring.Build(template, new[] { source }, directory.Path, output, producers: producers);
        Assert.Contains(strict, diagnostic => diagnostic.Code == "YADG-PRODUCER-012" && diagnostic.IsError);
        Assert.Equal("committed sentinel", File.ReadAllText(output));

        var yolo = PresentationAuthoring.Build(template, new[] { source }, directory.Path, output, yolo: true, producers: producers);
        Assert.DoesNotContain(yolo, diagnostic => diagnostic.IsError);
        Assert.Single(yolo, diagnostic => diagnostic.Code == "YADG-YOLO-PRODUCER-001" && diagnostic.IsDegradation);
        using var presentation = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(output, false);
        Assert.Contains(presentation.PresentationPart!.SlideParts.SelectMany(slide => slide.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>()), text => text.Text?.Contains("YADG image unavailable", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void Successful_mermaid_output_remains_owned_until_caller_cleanup()
    {
        using var directory = new TemporaryDirectory();
        var script = Path.Combine(directory.Path, "valid.ps1");
        File.WriteAllText(script, "[IO.File]::WriteAllBytes($args[3], [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='))");
        var ownedDirectories = new List<string>();
        var diagnostics = new List<Diagnostic>();
        var asset = MermaidProducer.RenderSource("figure", "flowchart LR\nA-->B", Path.Combine(directory.Path, "source.md"), new("powershell.exe", new[] { "-NoProfile", "-File", script }), directory.Path, diagnostics, ownedDirectories);

        Assert.NotNull(asset);
        var ownedDirectory = Assert.Single(ownedDirectories);
        Assert.True(Directory.Exists(ownedDirectory));
        Assert.True(File.Exists(asset.FullPath));
        Directory.Delete(ownedDirectory, true);
        Assert.False(Directory.Exists(ownedDirectory));
        Assert.Empty(diagnostics);
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

    private static void AssertFailedProducerIsCleaned(string expectedDiagnostic, MermaidProducerConfiguration configuration, string workspaceRoot, int? timeoutMilliseconds = null)
    {
        var temporaryDirectories = new List<string>();
        var diagnostics = new List<Diagnostic>();
        var sourcePath = Path.Combine(workspaceRoot, "source.md");
        var image = timeoutMilliseconds is int timeout
            ? MermaidProducer.RenderSource("cleanup-test", "flowchart LR\nA-->B", sourcePath, configuration, workspaceRoot, diagnostics, temporaryDirectories, timeout)
            : MermaidProducer.RenderSource("cleanup-test", "flowchart LR\nA-->B", sourcePath, configuration, workspaceRoot, diagnostics, temporaryDirectories);
        Assert.Null(image);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == expectedDiagnostic && diagnostic.IsError);
        var ownedDirectory = Assert.Single(temporaryDirectories);
        Assert.False(Directory.Exists(ownedDirectory), $"Failed producer left temporary directory '{ownedDirectory}'.");
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
