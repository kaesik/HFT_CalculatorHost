using System.IO;

namespace CalculatorHost.Services;

public class WorkingCopyService : IDisposable {
    private readonly string _baseDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CalculatorHost", "Sessions");

    private string? _currentSessionDirectory;

    public void Dispose() {
        CleanCurrentSession();
    }

    public string CreateWorkingCopy(string sourceFilePath) {
        CleanCurrentSession();

        var sessionId = Guid.NewGuid().ToString("N");
        _currentSessionDirectory = Path.Combine(_baseDirectory, sessionId);
        Directory.CreateDirectory(_currentSessionDirectory);

        var fileName = Path.GetFileName(sourceFilePath);
        var destinationPath = Path.Combine(_currentSessionDirectory, fileName);

        File.Copy(sourceFilePath, destinationPath, true);

        // Remove read-only attribute if present
        File.SetAttributes(destinationPath, FileAttributes.Normal);

        return destinationPath;
    }

    public void CleanCurrentSession() {
        if (_currentSessionDirectory == null) return;
        if (!Directory.Exists(_currentSessionDirectory)) return;

        try {
            // Attempt to set all files as normal before deleting (remove read-only)
            foreach (var file in Directory.GetFiles(_currentSessionDirectory))
                File.SetAttributes(file, FileAttributes.Normal);

            Directory.Delete(_currentSessionDirectory, true);
        }
        catch {
            // Best-effort cleanup; leftover session folders are cleaned on next startup
        }
        finally {
            _currentSessionDirectory = null;
        }
    }

    public void CleanAllOrphanedSessions() {
        if (!Directory.Exists(_baseDirectory)) return;

        foreach (var directory in Directory.GetDirectories(_baseDirectory))
            try {
                foreach (var file in Directory.GetFiles(directory))
                    File.SetAttributes(file, FileAttributes.Normal);
                Directory.Delete(directory, true);
            }
            catch {
                // Ignore directories that cannot be deleted (e.g., still locked by another process)
            }
    }
}