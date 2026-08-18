using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public class SheetLayoutCacheService {
    private const int CurrentCacheVersion = 1;

    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly string _cacheDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CalculatorHost",
        "SheetLayoutCache");

    public bool TryLoad(CalculatorInfo calculatorInfo, out SheetModel? model) {
        model = null;

        try {
            if (!File.Exists(calculatorInfo.FilePath))
                return false;

            var exactCachePath = GetCachePath(calculatorInfo.FilePath);

            if (TryLoadFromCacheFile(
                    exactCachePath,
                    calculatorInfo,
                    true,
                    out model))
                return true;

            if (TryLoadFromCacheFile(
                    exactCachePath,
                    calculatorInfo,
                    false,
                    out model))
                return true;

            return TryLoadCompatibleCache(calculatorInfo, exactCachePath, out model);
        }
        catch {
            model = null;
            return false;
        }
    }

    public bool IsCacheCurrent(CalculatorInfo calculatorInfo) {
        try {
            if (!File.Exists(calculatorInfo.FilePath))
                return false;

            var cachePath = GetCachePath(calculatorInfo.FilePath);
            if (!File.Exists(cachePath))
                return false;

            var json = File.ReadAllText(cachePath);
            var document = JsonSerializer.Deserialize<SheetLayoutCacheDocument>(json, JsonOptions);

            return document != null &&
                   document.Version == CurrentCacheVersion &&
                   document.Sheet != null &&
                   IsCurrentFileCache(document, calculatorInfo);
        }
        catch {
            return false;
        }
    }

    public bool TrySave(CalculatorInfo calculatorInfo, SheetModel model) {
        string? temporaryPath = null;

        try {
            if (!File.Exists(calculatorInfo.FilePath))
                return false;

            Directory.CreateDirectory(_cacheDirectory);

            var sourceFileInformation = new FileInfo(calculatorInfo.FilePath);
            var document = new SheetLayoutCacheDocument {
                Version = CurrentCacheVersion,
                SourceFileName = GetCalculatorFileName(calculatorInfo),
                CalculatorName = calculatorInfo.DisplayName,
                SourceFileLength = sourceFileInformation.Length,
                SourceLastWriteTimeUtcTicks = sourceFileInformation.LastWriteTimeUtc.Ticks,
                Sheet = ConvertToCachedSheet(model)
            };

            var cachePath = GetCachePath(calculatorInfo.FilePath);
            temporaryPath = $"{cachePath}.{Guid.NewGuid():N}.temporary";
            var json = JsonSerializer.Serialize(document, JsonOptions);

            File.WriteAllText(temporaryPath, json);
            File.Move(temporaryPath, cachePath, true);

            return true;
        }
        catch {
            if (!string.IsNullOrWhiteSpace(temporaryPath))
                try {
                    if (File.Exists(temporaryPath))
                        File.Delete(temporaryPath);
                }
                catch {
                    // ignored
                }

            return false;
        }
    }

    private bool TryLoadCompatibleCache(
        CalculatorInfo calculatorInfo,
        string exactCachePath,
        out SheetModel? model) {
        model = null;

        if (!Directory.Exists(_cacheDirectory))
            return false;

        foreach (var cachePath in Directory
                     .EnumerateFiles(_cacheDirectory, "*.json")
                     .Where(path => !string.Equals(path, exactCachePath, StringComparison.OrdinalIgnoreCase))
                     .OrderByDescending(File.GetLastWriteTimeUtc))
            if (TryLoadFromCacheFile(
                    cachePath,
                    calculatorInfo,
                    false,
                    out model))
                return true;

        return false;
    }

    private static bool TryLoadFromCacheFile(
        string cachePath,
        CalculatorInfo calculatorInfo,
        bool requireCurrentFileMetadata,
        out SheetModel? model) {
        model = null;

        try {
            if (!File.Exists(cachePath))
                return false;

            var json = File.ReadAllText(cachePath);
            var document = JsonSerializer.Deserialize<SheetLayoutCacheDocument>(json, JsonOptions);

            if (document == null ||
                document.Version != CurrentCacheVersion ||
                document.Sheet == null)
                return false;

            if (requireCurrentFileMetadata) {
                if (!IsCurrentFileCache(document, calculatorInfo))
                    return false;
            }
            else if (!IsCompatibleCalculatorCache(document, calculatorInfo)) return false;

            model = ConvertToSheetModel(document.Sheet);
            return true;
        }
        catch {
            model = null;
            return false;
        }
    }

    private static bool IsCurrentFileCache(
        SheetLayoutCacheDocument document,
        CalculatorInfo calculatorInfo) {
        if (!File.Exists(calculatorInfo.FilePath))
            return false;

        var sourceFileInformation = new FileInfo(calculatorInfo.FilePath);

        return document.SourceFileLength == sourceFileInformation.Length &&
               document.SourceLastWriteTimeUtcTicks == sourceFileInformation.LastWriteTimeUtc.Ticks &&
               document.Sheet != null;
    }

    private static bool IsCompatibleCalculatorCache(
        SheetLayoutCacheDocument document,
        CalculatorInfo calculatorInfo) {
        if (document.Sheet == null)
            return false;

        var currentFileName = GetCalculatorFileName(calculatorInfo);

        if (!string.IsNullOrWhiteSpace(currentFileName) &&
            !string.IsNullOrWhiteSpace(document.SourceFileName) &&
            string.Equals(currentFileName, document.SourceFileName, StringComparison.OrdinalIgnoreCase))
            return true;

        return !string.IsNullOrWhiteSpace(document.CalculatorName) &&
               string.Equals(document.CalculatorName, calculatorInfo.DisplayName, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetCalculatorFileName(CalculatorInfo calculatorInfo) {
        if (!string.IsNullOrWhiteSpace(calculatorInfo.FileName))
            return calculatorInfo.FileName;

        return Path.GetFileName(calculatorInfo.FilePath);
    }

    private string GetCachePath(string sourceFilePath) {
        var normalizedPath = Path.GetFullPath(sourceFilePath).ToUpperInvariant();
        var bytes = Encoding.UTF8.GetBytes(normalizedPath);
        var hash = Convert.ToHexString(SHA256.HashData(bytes));

        return Path.Combine(_cacheDirectory, $"{hash}.json");
    }

    private static CachedSheetModel ConvertToCachedSheet(SheetModel model) {
        return new CachedSheetModel {
            SheetName = model.SheetName,
            FirstRow = model.FirstRow,
            FirstColumn = model.FirstColumn,
            MaxRow = model.MaxRow,
            MaxColumn = model.MaxColumn,
            DefaultColumnWidth = model.DefaultColumnWidth,
            DefaultRowHeight = model.DefaultRowHeight,
            ColumnWidths = new Dictionary<int, double>(model.ColumnWidths),
            RowHeights = new Dictionary<int, double>(model.RowHeights),
            Cells = model.Cells.Select(ConvertToCachedCell).ToList()
        };
    }

    private static CachedCellModel ConvertToCachedCell(CellModel cell) {
        return new CachedCellModel {
            Row = cell.Row,
            Column = cell.Column,
            RowSpan = cell.RowSpan,
            ColumnSpan = cell.ColSpan,
            IsMergedSlave = cell.IsMergedSlave,
            BackgroundColor = ConvertColorToUnsignedInteger(cell.BackgroundColor),
            ForegroundColor = ConvertColorToUnsignedInteger(cell.ForegroundColor),
            IsBold = cell.IsBold,
            IsItalic = cell.IsItalic,
            FontSize = cell.FontSize,
            TextAlignment = cell.TextAlignment,
            VerticalContentAlignment = cell.VerticalContentAlignment,
            WrapText = cell.WrapText,
            BorderTopThickness = cell.BorderTopThickness,
            BorderBottomThickness = cell.BorderBottomThickness,
            BorderLeftThickness = cell.BorderLeftThickness,
            BorderRightThickness = cell.BorderRightThickness,
            BorderColor = ConvertColorToUnsignedInteger(cell.BorderColor),
            IsInput = cell.IsInput,
            InputType = cell.InputType,
            DropdownValues = [.. cell.DropdownValues]
        };
    }

    private static SheetModel ConvertToSheetModel(CachedSheetModel cachedSheet) {
        return new SheetModel {
            SheetName = cachedSheet.SheetName,
            FirstRow = cachedSheet.FirstRow,
            FirstColumn = cachedSheet.FirstColumn,
            MaxRow = cachedSheet.MaxRow,
            MaxColumn = cachedSheet.MaxColumn,
            DefaultColumnWidth = cachedSheet.DefaultColumnWidth,
            DefaultRowHeight = cachedSheet.DefaultRowHeight,
            ColumnWidths = new Dictionary<int, double>(cachedSheet.ColumnWidths),
            RowHeights = new Dictionary<int, double>(cachedSheet.RowHeights),
            Cells = cachedSheet.Cells.Select(ConvertToCellModel).ToList()
        };
    }

    private static CellModel ConvertToCellModel(CachedCellModel cachedCell) {
        return new CellModel {
            Row = cachedCell.Row,
            Column = cachedCell.Column,
            RowSpan = cachedCell.RowSpan,
            ColSpan = cachedCell.ColumnSpan,
            IsMergedSlave = cachedCell.IsMergedSlave,
            BackgroundColor = ConvertUnsignedIntegerToColor(cachedCell.BackgroundColor),
            ForegroundColor = ConvertUnsignedIntegerToColor(cachedCell.ForegroundColor),
            IsBold = cachedCell.IsBold,
            IsItalic = cachedCell.IsItalic,
            FontSize = cachedCell.FontSize,
            TextAlignment = cachedCell.TextAlignment,
            VerticalContentAlignment = cachedCell.VerticalContentAlignment,
            WrapText = cachedCell.WrapText,
            BorderTopThickness = cachedCell.BorderTopThickness,
            BorderBottomThickness = cachedCell.BorderBottomThickness,
            BorderLeftThickness = cachedCell.BorderLeftThickness,
            BorderRightThickness = cachedCell.BorderRightThickness,
            BorderColor = ConvertUnsignedIntegerToColor(cachedCell.BorderColor),
            IsInput = cachedCell.IsInput,
            InputType = cachedCell.InputType,
            DropdownValues = [.. cachedCell.DropdownValues]
        };
    }

    private static uint ConvertColorToUnsignedInteger(Color color) {
        return ((uint)color.A << 24) |
               ((uint)color.R << 16) |
               ((uint)color.G << 8) |
               color.B;
    }

    private static Color ConvertUnsignedIntegerToColor(uint color) {
        return Color.FromArgb(
            (byte)(color >> 24),
            (byte)(color >> 16),
            (byte)(color >> 8),
            (byte)color);
    }
}

public sealed class SheetLayoutCacheDocument {
    public int Version { get; init; }
    public string SourceFileName { get; init; } = string.Empty;
    public string CalculatorName { get; init; } = string.Empty;
    public long SourceFileLength { get; init; }
    public long SourceLastWriteTimeUtcTicks { get; init; }
    public CachedSheetModel? Sheet { get; init; }
}

public sealed class CachedSheetModel {
    public string SheetName { get; init; } = string.Empty;
    public List<CachedCellModel> Cells { get; init; } = [];
    public Dictionary<int, double> ColumnWidths { get; init; } = new();
    public Dictionary<int, double> RowHeights { get; init; } = new();
    public int FirstRow { get; init; } = 1;
    public int FirstColumn { get; init; } = 1;
    public int MaxRow { get; init; }
    public int MaxColumn { get; init; }
    public double DefaultColumnWidth { get; init; } = 64.0;
    public double DefaultRowHeight { get; init; } = 20.0;
}

public sealed class CachedCellModel {
    public int Row { get; init; }
    public int Column { get; init; }
    public int RowSpan { get; init; } = 1;
    public int ColumnSpan { get; init; } = 1;
    public bool IsMergedSlave { get; init; }
    public uint BackgroundColor { get; init; }
    public uint ForegroundColor { get; init; }
    public bool IsBold { get; init; }
    public bool IsItalic { get; init; }
    public double FontSize { get; init; } = 11.0;
    public TextAlignment TextAlignment { get; init; } = TextAlignment.Left;
    public VerticalAlignment VerticalContentAlignment { get; init; } = VerticalAlignment.Center;
    public bool WrapText { get; init; }
    public double BorderTopThickness { get; init; }
    public double BorderBottomThickness { get; init; }
    public double BorderLeftThickness { get; init; }
    public double BorderRightThickness { get; init; }
    public uint BorderColor { get; init; }
    public bool IsInput { get; init; }
    public CellInputType InputType { get; init; }
    public List<string> DropdownValues { get; init; } = [];
}