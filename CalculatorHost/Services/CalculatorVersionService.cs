using System.IO;
using System.Text.Json;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public static class CalculatorVersionService {
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

        if (version.FormatVersion != 1)
            throw new InvalidOperationException($"Nieobsługiwana wersja formatu pliku: {version.FormatVersion}.");

        return version;
    }
}
