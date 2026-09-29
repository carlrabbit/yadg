namespace Yadg.Renderer;

public static class FinalizedOutputCommit
{
    public static void Commit(string workspaceRoot, string stagedWords, string? stagedPdfs)
    {
        var root = Path.GetFullPath(workspaceRoot);
        var transaction = Path.Combine(root, ".yadg-render-commit-" + Guid.NewGuid().ToString("N"));
        var words = Path.Combine(root, "YadgWords");
        var pdfs = Path.Combine(root, "YadgPdfs");
        var oldWords = Path.Combine(transaction, "previous-words");
        var oldPdfs = Path.Combine(transaction, "previous-pdfs");
        var movedOldWords = false; var movedOldPdfs = false;
        var installedWords = false; var installedPdfs = false;
        try
        {
            var nextWords = Path.Combine(transaction, "YadgWords");
            var nextPdfs = Path.Combine(transaction, "YadgPdfs");
            Directory.CreateDirectory(nextWords);
            CopyFiles(stagedWords, nextWords);
            Directory.CreateDirectory(nextPdfs);
            if (stagedPdfs is not null) CopyFiles(stagedPdfs, nextPdfs);

            if (Directory.Exists(words)) { Directory.Move(words, oldWords); movedOldWords = true; }
            if (Directory.Exists(pdfs)) { Directory.Move(pdfs, oldPdfs); movedOldPdfs = true; }
            Directory.Move(nextWords, words); installedWords = true;
            Directory.Move(nextPdfs, pdfs); installedPdfs = true;
        }
        catch
        {
            if (installedWords && Directory.Exists(words)) Directory.Delete(words, true);
            if (installedPdfs && Directory.Exists(pdfs)) Directory.Delete(pdfs, true);
            if (movedOldWords && Directory.Exists(oldWords)) Directory.Move(oldWords, words);
            if (movedOldPdfs && Directory.Exists(oldPdfs)) Directory.Move(oldPdfs, pdfs);
            throw;
        }
        finally { try { if (Directory.Exists(transaction)) Directory.Delete(transaction, true); } catch { } }
    }

    private static void CopyFiles(string source, string destination)
    {
        if (!Directory.Exists(source)) return;
        foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.TopDirectoryOnly)) File.Copy(file, Path.Combine(destination, Path.GetFileName(file)), true);
    }
}
