using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Yadg.Renderer;

namespace Yadg.WordRenderer;

[SupportedOSPlatform("windows")]
public sealed class WordRenderer
{
    private const int TimeoutMs = 120_000;

    public RenderResult Render(string workspaceRoot)
    {
        var diagnostics = new List<RendererDiagnostic>();
        var root = Path.GetFullPath(workspaceRoot);
        var preWords = Path.Combine(root, "YadgPreWords");
        if (!Directory.Exists(root)) return Fail(diagnostics, "YADG-WORD-001", $"Workspace directory does not exist: '{root}'.");
        if (!Directory.Exists(preWords)) return Fail(diagnostics, "YADG-WORD-002", $"Missing authored-input directory '{preWords}'.");
        var inputs = Directory.EnumerateFiles(preWords, "*.docx", SearchOption.TopDirectoryOnly).OrderBy(p => p, StringComparer.Ordinal).ToArray();
        if (inputs.Length == 0) return Fail(diagnostics, "YADG-WORD-003", $"No top-level DOCX inputs were found in '{preWords}'.");

        var session = Path.Combine(Path.GetTempPath(), "yadg-word-" + Guid.NewGuid().ToString("N"));
        var stage = Path.Combine(session, "stage");
        Directory.CreateDirectory(stage);
        var result = new RenderResult(false, diagnostics);
        var thread = new Thread(() => result = RenderOnSta(root, inputs, stage, diagnostics)) { IsBackground = true, Name = "YADG Microsoft Word renderer" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        if (!thread.Join(TimeoutMs))
        {
            diagnostics.Add(new("YADG-WORD-010", $"Microsoft Word rendering exceeded the {TimeoutMs / 1000}-second bound; the owned automation thread did not complete."));
            result = new(false, diagnostics);
        }
        try { if (Directory.Exists(session)) Directory.Delete(session, true); } catch { }
        return result;
    }

    private static RenderResult RenderOnSta(string root, IReadOnlyList<string> inputs, string stage, List<RendererDiagnostic> diagnostics)
    {
        object? app = null;
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application", throwOnError: true)
                ?? throw new InvalidOperationException("Word.Application is not registered.");
            dynamic word = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Word.Application could not be activated.");
            app = word;
            word.Visible = false;
            word.DisplayAlerts = 0;
            word.ScreenUpdating = false;
            var version = (string)word.Version;
            try { word.AutomationSecurity = 3; } catch { }
            foreach (var input in inputs)
            {
                var stagedInput = Path.Combine(stage, "input-" + Path.GetFileName(input));
                var stagedOutput = Path.Combine(stage, Path.GetFileName(input));
                File.Copy(input, stagedInput, true);
                object? document = null;
                try
                {
                    dynamic documents = word.Documents;
                    dynamic opened = documents.Open(FileName: stagedInput, ReadOnly: false, AddToRecentFiles: false, Visible: false, OpenAndRepair: false);
                    document = opened;
                    UpdateDocument(document);
                    ((dynamic)document).SaveAs2(FileName: stagedOutput, FileFormat: 16, AddToRecentFiles: false);
                }
                catch (Exception ex)
                {
                    diagnostics.Add(new("YADG-WORD-011", $"Word finalization failed at open/refresh/save for '{input}': {ex.Message}", input));
                    return new(false, diagnostics, version);
                }
                finally
                {
                    if (document is not null) try { ((dynamic)document).Close(0); } catch { }
                    Release(document);
                }
                if (!File.Exists(stagedOutput) || new FileInfo(stagedOutput).Length == 0)
                    return Fail(diagnostics, "YADG-WORD-012", $"Word did not produce a finalized DOCX for '{input}'.", version);
            }
            var output = Path.Combine(root, "YadgWords");
            Directory.CreateDirectory(output);
            foreach (var input in inputs)
                File.Move(Path.Combine(stage, Path.GetFileName(input)), Path.Combine(output, Path.GetFileName(input)), true);
            return new(true, diagnostics, version);
        }
        catch (COMException ex)
        {
            diagnostics.Add(new("YADG-WORD-004", $"Microsoft Word COM activation or interrogation failed (0x{ex.HResult:X8}): {ex.Message}"));
            return new(false, diagnostics);
        }
        catch (Exception ex)
        {
            diagnostics.Add(new("YADG-WORD-005", $"Microsoft Word renderer failed: {ex.Message}"));
            return new(false, diagnostics);
        }
        finally
        {
            if (app is not null) try { ((dynamic)app).Quit(0); } catch { }
            Release(app);
        }
    }

    private static void UpdateDocument(object document)
    {
        dynamic doc = document;
        doc.Repaginate();
        foreach (dynamic field in doc.Fields) field.Update();
        foreach (dynamic toc in doc.TablesOfContents) toc.Update();
        foreach (dynamic tof in doc.TablesOfFigures) tof.Update();
        foreach (dynamic section in doc.Sections)
        {
            foreach (dynamic header in section.Headers) foreach (dynamic field in header.Range.Fields) field.Update();
            foreach (dynamic footer in section.Footers) foreach (dynamic field in footer.Range.Fields) field.Update();
        }
        doc.Repaginate();
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) try { Marshal.FinalReleaseComObject(value); } catch { }
    }

    private static RenderResult Fail(List<RendererDiagnostic> diagnostics, string code, string message, string? version = null)
    {
        diagnostics.Add(new(code, message));
        return new(false, diagnostics, version);
    }
}
