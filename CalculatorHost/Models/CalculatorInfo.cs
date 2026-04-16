using System.IO;

namespace CalculatorHost.Models;

public class CalculatorInfo {
    public string FilePath { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public string DisplayName => Path.GetFileNameWithoutExtension(FileName);
    public DateTime LastModified { get; init; }
    public long FileSizeBytes { get; init; }

    public string FileSizeDisplay => FileSizeBytes < 1_048_576
        ? $"{FileSizeBytes / 1024.0:N1} KB"
        : $"{FileSizeBytes / 1_048_576.0:N1} MB";

    public string LastModifiedDisplay => LastModified.ToString("yyyy-MM-dd HH:mm");
}