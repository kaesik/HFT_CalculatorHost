using System.IO;
using System.Security.Cryptography;
using System.Text.Json;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public static class CalculatorVersionService {
    public const int CurrentFormatVersion = 2;

    private const string ApplicationDirectoryName = "CalculatorHost";
    private const string VersionsDirectoryName = "Versions";

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

        return version.FormatVersion is < 1 or > CurrentFormatVersion
            ? throw new InvalidOperationException($"Nieobsługiwana wersja formatu pliku: {version.FormatVersion}.")
            : version;
    }

    public static string GetVersionsDirectory(string calculatorFilePath) {
        var calculatorName = Path.GetFileNameWithoutExtension(calculatorFilePath);
        var safeCalculatorName = MakeSafeFileName(calculatorName);

        return Path.Combine(
            GetVersionsRootDirectory(),
            safeCalculatorName);
    }

    public static string CreateDefaultVersionFilePath(string calculatorFilePath) {
        var directory = GetVersionsDirectory(calculatorFilePath);
        Directory.CreateDirectory(directory);

        return Path.Combine(directory, $"wersja_{DateTime.Now:yyyyMMdd_HHmmss}.json");
    }

    public static List<CalculatorVersionFileInfo> FindMatchingVersionFiles(string calculatorFilePath) {
        var currentIdentity = CreateFileIdentity(calculatorFilePath);
        var results = new List<CalculatorVersionFileInfo>();
        var visitedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var directory in EnumerateCandidateVersionDirectories(calculatorFilePath)) {
            if (!Directory.Exists(directory))
                continue;

            foreach (var filePath in Directory.EnumerateFiles(directory, "*.json", SearchOption.TopDirectoryOnly)) {
                if (!visitedFiles.Add(filePath))
                    continue;

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
                        ValuesCount = version.Values.Count,
                        IsExactFileMatch = IsExactIdentityMatch(version.CalculatorFileIdentity, currentIdentity)
                    });
                }
                catch {
                    // Uszkodzony albo obcy JSON nie powinien blokować listy wersji.
                }
            }
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

        if (savedIdentity == null)
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

    private static string GetVersionsRootDirectory() {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            ApplicationDirectoryName,
            VersionsDirectoryName);
    }

    private static IEnumerable<string> EnumerateCandidateVersionDirectories(string calculatorFilePath) {
        var versionsRootDirectory = GetVersionsRootDirectory();
        var currentDirectory = GetVersionsDirectory(calculatorFilePath);

        yield return currentDirectory;

        if (!Directory.Exists(versionsRootDirectory))
            yield break;

        var calculatorName = Path.GetFileNameWithoutExtension(calculatorFilePath);
        var safeCalculatorName = MakeSafeFileName(calculatorName);
        var legacyDirectoryPrefix = safeCalculatorName + "_";

        foreach (var directory in Directory.EnumerateDirectories(versionsRootDirectory, legacyDirectoryPrefix + "*",
                     SearchOption.TopDirectoryOnly)) {
            var directoryName = Path.GetFileName(directory);

            if (directoryName.StartsWith(legacyDirectoryPrefix, StringComparison.OrdinalIgnoreCase))
                yield return directory;
        }
    }

    private static bool BelongsToCalculator(
        CalculatorVersionModel version,
        CalculatorVersionFileIdentityModel currentIdentity) {
        var savedIdentity = version.CalculatorFileIdentity;

        if (savedIdentity == null)
            return false;

        if (IsExactIdentityMatch(savedIdentity, currentIdentity))
            return true;

        return IsOlderVersionOfSameCalculator(version, savedIdentity, currentIdentity);
    }

    private static bool IsExactIdentityMatch(
        CalculatorVersionFileIdentityModel? savedIdentity,
        CalculatorVersionFileIdentityModel currentIdentity) {
        return savedIdentity != null &&
               !string.IsNullOrWhiteSpace(savedIdentity.Sha256) &&
               string.Equals(savedIdentity.Sha256, currentIdentity.Sha256, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsOlderVersionOfSameCalculator(
        CalculatorVersionModel version,
        CalculatorVersionFileIdentityModel savedIdentity,
        CalculatorVersionFileIdentityModel currentIdentity) {
        if (string.IsNullOrWhiteSpace(savedIdentity.FileName) ||
            !string.Equals(savedIdentity.FileName, currentIdentity.FileName, StringComparison.OrdinalIgnoreCase))
            return false;

        var currentCalculatorName = Path.GetFileNameWithoutExtension(currentIdentity.FileName);

        return string.IsNullOrWhiteSpace(version.CalculatorName) ||
               string.Equals(version.CalculatorName, currentCalculatorName, StringComparison.OrdinalIgnoreCase);
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
    public bool IsExactFileMatch { get; init; }

    public string DisplayName {
        get {
            var matchInfo = IsExactFileMatch ? string.Empty : " · nieaktualny";
            return
                $"{Path.GetFileNameWithoutExtension(FileName)} · {CreatedAt:yyyy-MM-dd HH:mm} · {ValuesCount} pól{matchInfo}";
        }
    }
}