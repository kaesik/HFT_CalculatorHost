using System.IO;
using System.Text.Json;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public sealed class UserPreferencesService {
    private const string ApplicationDirectoryName = "CalculatorHost";
    private const string PreferencesFileName = "calculator-list-preferences.json";

    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    private readonly object _syncRoot = new();
    private readonly CalculatorUserPreferences _preferences;

    public UserPreferencesService() {
        _preferences = Load();
    }

    public string SortMode {
        get {
            lock (_syncRoot) {
                return _preferences.SortMode;
            }
        }
    }

    public string FilterMode {
        get {
            lock (_syncRoot) {
                return _preferences.FilterMode;
            }
        }
    }

    private static string PreferencesFilePath {
        get {
            var localApplicationData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var directory = string.IsNullOrWhiteSpace(localApplicationData)
                ? AppDomain.CurrentDomain.BaseDirectory
                : Path.Combine(localApplicationData, ApplicationDirectoryName);

            return Path.Combine(directory, PreferencesFileName);
        }
    }

    public bool IsFavorite(string filePath) {
        if (string.IsNullOrWhiteSpace(filePath))
            return false;

        lock (_syncRoot) {
            return _preferences.FavoriteCalculatorPaths.Contains(NormalizePath(filePath));
        }
    }

    public void SetFavorite(string filePath, bool isFavorite) {
        if (string.IsNullOrWhiteSpace(filePath))
            return;

        lock (_syncRoot) {
            var normalizedPath = NormalizePath(filePath);

            if (isFavorite)
                _preferences.FavoriteCalculatorPaths.Add(normalizedPath);
            else
                _preferences.FavoriteCalculatorPaths.Remove(normalizedPath);

            Save();
        }
    }

    public DateTime? GetLastOpenedUtc(string filePath) {
        if (string.IsNullOrWhiteSpace(filePath))
            return null;

        lock (_syncRoot) {
            return _preferences.LastOpenedUtcByPath.TryGetValue(NormalizePath(filePath), out var openedAt)
                ? openedAt
                : null;
        }
    }

    public DateTime MarkOpened(string filePath) {
        var openedAt = DateTime.UtcNow;
        if (string.IsNullOrWhiteSpace(filePath))
            return openedAt;

        lock (_syncRoot) {
            _preferences.LastOpenedUtcByPath[NormalizePath(filePath)] = openedAt;
            Save();
        }

        return openedAt;
    }

    public void SetListOptions(string sortMode, string filterMode) {
        lock (_syncRoot) {
            _preferences.SortMode = CalculatorListOptions.IsValidSortMode(sortMode)
                ? sortMode
                : CalculatorListOptions.SortNameAscending;

            _preferences.FilterMode = CalculatorListOptions.IsValidFilterMode(filterMode)
                ? filterMode
                : CalculatorListOptions.FilterAll;

            _preferences.FavoritesOnly = _preferences.FilterMode == CalculatorListOptions.FilterFavorites;
            Save();
        }
    }

    private static CalculatorUserPreferences Load() {
        try {
            if (!File.Exists(PreferencesFilePath))
                return new CalculatorUserPreferences();

            var json = File.ReadAllText(PreferencesFilePath);
            var loaded = JsonSerializer.Deserialize<CalculatorUserPreferences>(json, JsonOptions)
                         ?? new CalculatorUserPreferences();

            loaded.FavoriteCalculatorPaths = new HashSet<string>(
                loaded.FavoriteCalculatorPaths ?? [],
                StringComparer.OrdinalIgnoreCase);

            loaded.LastOpenedUtcByPath = new Dictionary<string, DateTime>(
                loaded.LastOpenedUtcByPath ?? [],
                StringComparer.OrdinalIgnoreCase);

            if (!CalculatorListOptions.IsValidSortMode(loaded.SortMode))
                loaded.SortMode = CalculatorListOptions.SortNameAscending;

            if (!CalculatorListOptions.IsValidGroupMode(loaded.GroupMode))
                loaded.GroupMode = CalculatorListOptions.GroupFavorites;

            if (loaded.FavoritesOnly && loaded.FilterMode == CalculatorListOptions.FilterAll)
                loaded.FilterMode = CalculatorListOptions.FilterFavorites;
            else if (!CalculatorListOptions.IsValidFilterMode(loaded.FilterMode))
                loaded.FilterMode = loaded.FavoritesOnly
                    ? CalculatorListOptions.FilterFavorites
                    : CalculatorListOptions.FilterAll;

            return loaded;
        }
        catch {
            // Uszkodzony plik preferencji nie powinien blokować uruchomienia programu.
            return new CalculatorUserPreferences();
        }
    }

    private void Save() {
        try {
            var filePath = PreferencesFilePath;
            var directory = Path.GetDirectoryName(filePath);

            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            var json = JsonSerializer.Serialize(_preferences, JsonOptions);
            var temporaryPath = filePath + ".tmp";

            File.WriteAllText(temporaryPath, json);
            File.Move(temporaryPath, filePath, true);
        }
        catch {
            // Preferencje są dodatkiem; błąd zapisu nie może zatrzymać kalkulatora.
        }
    }

    private static string NormalizePath(string filePath) {
        if (string.IsNullOrWhiteSpace(filePath))
            return string.Empty;

        try {
            return Path.GetFullPath(filePath.Trim());
        }
        catch {
            return filePath.Trim();
        }
    }
}