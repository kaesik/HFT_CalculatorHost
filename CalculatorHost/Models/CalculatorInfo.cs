using System.IO;

namespace CalculatorHost.Models;

public class CalculatorInfo {
    public string FilePath { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public string DisplayName => Path.GetFileNameWithoutExtension(FileName);
    public DateTime LastModified { get; init; }
    public long FileSizeBytes { get; init; }
}