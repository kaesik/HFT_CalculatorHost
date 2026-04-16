using System.IO;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public class CalculatorScannerService {
    // *** KONFIGURACJA: zmień tę ścieżkę, aby wskazać folder z kalkulatorami ***
    public string CalculatorsFolder { get; set; } = @"C:\Users\PMJ-KSobania2\Desktop\pliki\smieci\HFT_ExcelSite";

    public List<CalculatorInfo> ScanFolder() {
        if (!Directory.Exists(CalculatorsFolder))
            return [];

        var files = Directory.GetFiles(CalculatorsFolder, "*.xlsm", SearchOption.TopDirectoryOnly);
        var results = (from filePath in files
            let info = new FileInfo(filePath)
            select new CalculatorInfo {
                FilePath = filePath, FileName = info.Name, LastModified = info.LastWriteTime,
                FileSizeBytes = info.Length
            }).ToList();

        return results.OrderBy(c => c.DisplayName).ToList();
    }

    public bool FolderExists() {
        return Directory.Exists(CalculatorsFolder);
    }
}