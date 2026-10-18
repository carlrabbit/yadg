using Yadg.Renderer;

namespace Yadg.IntegrationTests;

public sealed class M0015RendererDiagnosticsTests
{
    [Fact]
    public void LibreOffice_timeout_reports_phase_and_drains_both_child_pipes()
    {
        var libreOffice = Environment.GetEnvironmentVariable("YADG_LIBREOFFICE_PATH") ?? @"C:\Program Files\LibreOffice\program\soffice.com";
        var python = Path.Combine(Path.GetDirectoryName(libreOffice)!, "python.exe");
        if (!File.Exists(python)) return; // The real Tier-4 workstation supplies LibreOffice's bundled UNO Python.

        var sidecar = Path.Combine(Path.GetTempPath(), "yadg-m0015-timeout-" + Guid.NewGuid().ToString("N") + ".py");
        try
        {
            File.WriteAllText(sidecar, "import json,sys,time\n" +
                "sys.stdout.write('o' * 262144 + '\\n'); sys.stdout.flush()\n" +
                "sys.stderr.write('e' * 262144 + '\\n'); sys.stderr.flush()\n" +
                "print(json.dumps({'type':'progress','phase':'document load','status':'started','elapsedMs':1}), flush=True)\n" +
                "time.sleep(30)\n");
            var diagnostics = new List<RendererDiagnostic>();
            var clock = System.Diagnostics.Stopwatch.StartNew();
            var success = LibreOfficeRenderer.RunSidecar(python, sidecar, 1, "input-review.docx", "out.docx", "out.pdf", diagnostics, null, "review.docx", 10000);
            var elapsed = clock.Elapsed;

            Assert.False(success);
            var timeout = Assert.Single(diagnostics);
            Assert.Equal("YADG-RENDER-009", timeout.Code);
            Assert.Contains("timeout=10000 ms", timeout.Message);
            Assert.Contains("elapsed=", timeout.Message);
            Assert.Contains("input='review.docx'", timeout.Message);
            Assert.Contains("current phase='document load'", timeout.Message);
            Assert.Contains("last completed phase='none'", timeout.Message);
            Assert.InRange(elapsed.TotalSeconds, 8.0, 18.0);
        }
        finally
        {
            try { File.Delete(sidecar); } catch { }
        }
    }
}
