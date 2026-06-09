using System.IO;
using System.Security.Cryptography;
using System.Text.Json;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public static class CalculatorVersionService {
    public const int CurrentFormatVersion = 2;

    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static void Save(string filePath, CalculatorVersionModel version) {
        var directory = Path.GetDirectoryName(filePath);

        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);

        var json = JsonSerializer.Serialize(version, JsonOptions);
        File.WriteAllText(filePath, json);
    }

    public static CalculatorVersionModel Load(string filePath) {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Nie znaleziono pliku wersji.", filePath);

        var json = File.ReadAllText(filePath);
        var version = JsonSerializer.Deserialize<CalculatorVersionModel>(json, JsonOptions);

        if (version == null)
            throw new InvalidOperationException("Plik wersji jest pusty albo ma niepoprawny format.");

        if (version.FormatVersion < 1 || version.FormatVersion > CurrentFormatVersion)
            throw new InvalidOperationException($"Nieobsługiwana wersja formatu pliku: {version.FormatVersion}.");

        return version;
    }

    public static string GetVersionsDirectory(string calculatorFilePath) {
        var identity = CreateFileIdentity(calculatorFilePath);
        var calculatorName = Path.GetFileNameWithoutExtension(calculatorFilePath);
        var safeCalculatorName = MakeSafeFileName(calculatorName);
        var shortHash = identity.Sha256.Length >= 16
            ? identity.Sha256[..16]
            : identity.Sha256;

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "CalculatorHost",
            "Versions",
            $"{safeCalculatorName}_{shortHash}");
    }

    public static string CreateDefaultVersionFilePath(string calculatorFilePath) {
        var directory = GetVersionsDirectory(calculatorFilePath);
        Directory.CreateDirectory(directory);

        return Path.Combine(directory, $"wersja_{DateTime.Now:yyyyMMdd_HHmmss}.json");
    }

    public static List<CalculatorVersionFileInfo> FindMatchingVersionFiles(string calculatorFilePath) {
        var directory = GetVersionsDirectory(calculatorFilePath);

        if (!Directory.Exists(directory))
            return [];

        var currentIdentity = CreateFileIdentity(calculatorFilePath);
        var results = new List<CalculatorVersionFileInfo>();

        foreach (var filePath in Directory.EnumerateFiles(directory, "*.json", SearchOption.TopDirectoryOnly))
            try {
                var version = Load(filePath);

                if (!BelongsToCalculator(version, currentIdentity))
                    continue;

                results.Add(new CalculatorVersionFileInfo {
                    FilePath = filePath,
                    FileName = Path.GetFileName(filePath),
                    CalculatorName = version.CalculatorName,
                    SheetName = version.SheetName,
                    CreatedAt = version.CreatedAt,
                    ValuesCount = version.Values.Count
                });
            }
            catch {
            }

        return results
            .OrderByDescending(file => file.CreatedAt)
            .ThenBy(file => file.FileName)
            .ToList();
    }

    public static CalculatorVersionFileIdentityModel CreateFileIdentity(string calculatorFilePath) {
        if (!File.Exists(calculatorFilePath))
            throw new FileNotFoundException("Nie znaleziono pliku kalkulatora.", calculatorFilePath);

        var fileInformation = new FileInfo(calculatorFilePath);

        using var sha256 = SHA256.Create();
        using var stream = File.OpenRead(calculatorFilePath);
        var hash = Convert.ToHexString(sha256.ComputeHash(stream));

        return new CalculatorVersionFileIdentityModel {
            FileName = fileInformation.Name,
            FileSizeBytes = fileInformation.Length,
            LastWriteTimeUtcTicks = fileInformation.LastWriteTimeUtc.Ticks,
            Sha256 = hash
        };
    }

    public static void ValidateBelongsToCalculator(
        CalculatorVersionModel version,
        string calculatorFilePath) {
        var savedIdentity = version.CalculatorFileIdentity;

        if (savedIdentity == null || string.IsNullOrWhiteSpace(savedIdentity.Sha256))
            throw new InvalidOperationException(
                "Ten plik wersji nie zawiera identyfikatora pliku kalkulatora. Zapisz wersję ponownie w aktualnej wersji programu.");

        var currentIdentity = CreateFileIdentity(calculatorFilePath);

        if (BelongsToCalculator(version, currentIdentity))
            return;

        throw new InvalidOperationException(
            $"Ten plik wersji należy do innego kalkulatora. " +
            $"Wersja jest zapisana dla pliku '{savedIdentity.FileName}', " +
            $"a aktualnie otwarty jest '{currentIdentity.FileName}'.");
    }

    private static bool BelongsToCalculator(
        CalculatorVersionModel version,
        CalculatorVersionFileIdentityModel currentIdentity) {
        var savedIdentity = version.CalculatorFileIdentity;

        if (savedIdentity == null || string.IsNullOrWhiteSpace(savedIdentity.Sha256))
            return false;

        return string.Equals(savedIdentity.Sha256, currentIdentity.Sha256, StringComparison.OrdinalIgnoreCase);
    }

    private static string MakeSafeFileName(string value) {
        var invalidCharacters = Path.GetInvalidFileNameChars().ToHashSet();
        var characters = value
            .Select(character => invalidCharacters.Contains(character) ? '_' : character)
            .ToArray();

        var result = new string(characters).Trim();
        return string.IsNullOrWhiteSpace(result) ? "calculator" : result;
    }
}

public class CalculatorVersionFileInfo {
    public string FilePath { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public string CalculatorName { get; init; } = string.Empty;
    public string SheetName { get; init; } = string.Empty;
    public DateTime CreatedAt { get; init; }
    public int ValuesCount { get; init; }

    public string DisplayName =>
        $"{Path.GetFileNameWithoutExtension(FileName)} · {CreatedAt:yyyy-MM-dd HH:mm} · {ValuesCount} pól";
}