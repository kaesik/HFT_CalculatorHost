namespace CalculatorHost.Models;

public sealed class CalculatorUserPreferences {
    public HashSet<string> FavoriteCalculatorPaths { get; set; } =
        new(StringComparer.OrdinalIgnoreCase);

    public Dictionary<string, DateTime> LastOpenedUtcByPath { get; set; } =
        new(StringComparer.OrdinalIgnoreCase);

    public string SortMode { get; set; } = CalculatorListOptions.SortNameAscending;
    public string FilterMode { get; set; } = CalculatorListOptions.FilterAll;

    // Pola zachowane dla zgodności z pierwszą wersją pliku preferencji.
    public string GroupMode { get; set; } = CalculatorListOptions.GroupFavorites;
    public bool FavoritesOnly { get; set; }
}

public static class CalculatorListOptions {
    public const string SortNameAscending = "Nazwa A-Z";
    public const string SortNameDescending = "Nazwa Z-A";
    public const string SortModifiedNewest = "Ostatnio zmienione";
    public const string SortRecentlyOpened = "Ostatnio używane";
    public const string SortFavoritesFirst = "Najpierw ulubione";

    public const string FilterAll = "Wszystkie";
    public const string FilterFavorites = "Ulubione";
    public const string FilterRecent = "Ostatnie";

    public const string GroupNone = "Bez grupowania";
    public const string GroupFavorites = "Ulubione";
    public const string GroupFirstLetter = "Pierwsza litera";

    public static IReadOnlyList<string> SortModes { get; } = [
        SortNameAscending,
        SortNameDescending,
        SortModifiedNewest,
        SortRecentlyOpened,
        SortFavoritesFirst
    ];

    public static IReadOnlyList<string> GroupModes { get; } = [
        GroupNone,
        GroupFavorites,
        GroupFirstLetter
    ];

    public static IReadOnlyList<string> FilterModes { get; } = [
        FilterAll,
        FilterFavorites,
        FilterRecent
    ];

    public static bool IsValidSortMode(string? value) {
        return !string.IsNullOrWhiteSpace(value) && SortModes.Contains(value);
    }

    public static bool IsValidGroupMode(string? value) {
        return !string.IsNullOrWhiteSpace(value) && GroupModes.Contains(value);
    }

    public static bool IsValidFilterMode(string? value) {
        return !string.IsNullOrWhiteSpace(value) && FilterModes.Contains(value);
    }
}