using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CalculatorHost.Models;
using CalculatorHost.Services;
using CalculatorHost.ViewModels;
using Microsoft.Win32;

namespace CalculatorHost.Views;

public partial class CalculatorListView {
    private readonly MainViewModel _viewModel;
    private List<CalculatorListEntry> _calculatorListItems = new();
    private bool _isRefreshingCalculatorList;
    private bool _isSubscribed;

    public CalculatorListView(MainViewModel viewModel) {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;

        RefreshCalculatorListItems();
        UpdateCount();
        UpdateSelectedCalculatorPanel();

        Loaded += CalculatorListView_Loaded;
        Unloaded += CalculatorListView_Unloaded;
    }

    private void CalculatorListView_Loaded(object sender, RoutedEventArgs e) {
        if (_isSubscribed) return;

        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
        _isSubscribed = true;

        RefreshCalculatorListItems();
        UpdateCount();
        UpdateSelectedCalculatorPanel();
    }

    private void CalculatorListView_Unloaded(object sender, RoutedEventArgs e) {
        if (!_isSubscribed) return;

        _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        _isSubscribed = false;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e) {
        switch (e.PropertyName) {
            case nameof(MainViewModel.Calculators):
                RefreshCalculatorListItems();
                UpdateCount();
                UpdateSelectedCalculatorPanel();
                break;
            case nameof(MainViewModel.SelectedCalculator):
                SelectCurrentCalculatorInList();
                UpdateSelectedCalculatorPanel();
                break;
        }
    }

    private void UpdateCount() {
        var count = _viewModel.Calculators.Count;
        CountText.Text = count switch {
            0 => "Brak kalkulatorów",
            1 => "1 kalkulator",
            _ => $"{count} kalkulatorów"
        };
    }

    private void RefreshCalculatorListItems() {
        _calculatorListItems = new List<CalculatorListEntry>();

        foreach (var calculator in _viewModel.Calculators)
            _calculatorListItems.Add(new CalculatorListEntry(calculator));

        _isRefreshingCalculatorList = true;
        try {
            CalculatorList.ItemsSource = _calculatorListItems;
            SelectCurrentCalculatorInList();
        }
        finally {
            _isRefreshingCalculatorList = false;
        }
    }

    private void SelectCurrentCalculatorInList() {
        var selectedCalculator = _viewModel.SelectedCalculator;

        if (selectedCalculator == null) {
            CalculatorList.SelectedItem = null;
            return;
        }

        foreach (var item in _calculatorListItems)
            if (string.Equals(item.FilePath, selectedCalculator.FilePath, StringComparison.OrdinalIgnoreCase)) {
                CalculatorList.SelectedItem = item;
                return;
            }
    }

    private void CalculatorList_SelectionChanged(object sender, SelectionChangedEventArgs e) {
        if (_isRefreshingCalculatorList) return;

        _viewModel.SelectedCalculator = CalculatorList.SelectedItem switch {
            CalculatorListEntry listItem => listItem.Calculator,
            CalculatorInfo calculator => calculator,
            _ => null
        };

        UpdateSelectedCalculatorPanel();
    }

    private void CalculatorList_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
        OpenCurrentCalculator();
    }

    private void CalculatorList_KeyDown(object sender, KeyEventArgs e) {
        if (e.Key != Key.Return && e.Key != Key.Enter) return;
        OpenCurrentCalculator();
    }

    private void VersionList_SelectionChanged(object sender, SelectionChangedEventArgs e) {
        // Otwieranie pozycji z listy zostaje pod dwuklikiem.
    }

    private void VersionList_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
        OpenSelectedVersion();
    }

    private void UploadVersionButton_Click(object sender, RoutedEventArgs e) {
        UploadVersionFromDisk();
    }

    private void FavoriteVersionButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if ((sender as FrameworkElement)?.Tag is not VersionListEntry version || version.IsCurrentFile)
            return;

        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        var newFavoriteState = !version.IsFavorite;
        VersionFavoritesStore.SetFavorite(version.FilePath, newFavoriteState);
        RefreshVersionListAndSelect(calculator, version.FilePath);

        VersionStatusText.Text = newFavoriteState
            ? $"Dodano do ulubionych: {version.DisplayName}."
            : $"Usunięto z ulubionych: {version.DisplayName}.";
    }

    private void DeleteVersionButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if ((sender as FrameworkElement)?.Tag is not VersionListEntry version || version.IsCurrentFile)
            return;

        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        var result = MessageBox.Show(
            Window.GetWindow(this)!,
            $"Usunąć wersję „{version.DisplayName}”?",
            "Usuń wersję",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
            return;

        try {
            VersionFavoritesStore.RemoveFavorite(version.FilePath);

            if (File.Exists(version.FilePath))
                File.Delete(version.FilePath);

            RefreshVersionListAndSelect(calculator, string.Empty);
            VersionStatusText.Text = $"Usunięto wersję: {version.DisplayName}.";
        }
        catch {
            MessageBox.Show(
                Window.GetWindow(this)!,
                "Nie udało się usunąć wybranej wersji.",
                "Nie udało się usunąć wersji",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void VersionActionButton_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e) {
        e.Handled = true;
    }

    private void UpdateSelectedCalculatorPanel() {
        var calculator = _viewModel.SelectedCalculator;

        if (calculator == null) {
            EmptySelectionPanel.Visibility = Visibility.Visible;
            CalculatorDetailsPanel.Visibility = Visibility.Collapsed;
            VersionList.ItemsSource = null;
            return;
        }

        EmptySelectionPanel.Visibility = Visibility.Collapsed;
        CalculatorDetailsPanel.Visibility = Visibility.Visible;

        SelectedCalculatorNameText.Text = calculator.DisplayName;
        SelectedCalculatorModifiedText.Text = FormatLastModified(calculator.FilePath);
        SelectedCalculatorSizeText.Text = FormatFileSize(calculator.FilePath);

        var versions = LoadVersionItems(calculator, out var versionStatusMessage);
        VersionList.ItemsSource = versions;
        VersionList.SelectedIndex = versions.Count > 0 ? 0 : -1;

        VersionStatusText.Text = versionStatusMessage;

        UpdateUploadVersionButton();
    }

    private List<VersionListEntry> LoadVersionItems(CalculatorInfo calculator, out string statusMessage) {
        var items = new List<VersionListEntry> {
            VersionListEntry.CreateCurrent(calculator)
        };

        try {
            var matchingVersions = CalculatorVersionService.FindMatchingVersionFiles(calculator.FilePath);
            var savedVersions = new List<VersionListEntry>();

            foreach (var matchingVersion in matchingVersions) {
                var filePath = GetStringProperty(matchingVersion, "FilePath")
                               ?? GetStringProperty(matchingVersion, "VersionFilePath")
                               ?? GetStringProperty(matchingVersion, "FullPath")
                               ?? GetStringProperty(matchingVersion, "Path")
                               ?? GetStringProperty(matchingVersion, "FullName");

                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    continue;

                var displayName = GetStringProperty(matchingVersion, "DisplayName")
                                  ?? GetStringProperty(matchingVersion, "FileName")
                                  ?? GetStringProperty(matchingVersion, "Name")
                                  ?? Path.GetFileNameWithoutExtension(filePath);

                var createdAt = GetDateTimeProperty(matchingVersion, "CreatedAt")
                                ?? GetDateTimeProperty(matchingVersion, "SavedAt")
                                ?? GetDateTimeProperty(matchingVersion, "ModifiedAt")
                                ?? GetDateTimeProperty(matchingVersion, "LastModified")
                                ?? GetDateTimeProperty(matchingVersion, "LastWriteTime")
                                ?? File.GetLastWriteTime(filePath);

                savedVersions.Add(VersionListEntry.CreateSavedVersion(
                    filePath,
                    displayName,
                    createdAt,
                    VersionFavoritesStore.IsFavorite(filePath)));
            }

            items.AddRange(savedVersions
                .OrderByDescending(version => version.IsFavorite)
                .ThenByDescending(version => version.CreatedAt)
                .ThenBy(version => version.DisplayName, StringComparer.CurrentCultureIgnoreCase));

            var favoriteCount = savedVersions.Count(version => version.IsFavorite);

            statusMessage = savedVersions.Count == 0
                ? "Brak zapisanych wersji pasujących do tego kalkulatora."
                : favoriteCount > 0
                    ? $"Znaleziono zapisane wersje: {savedVersions.Count}. Ulubione: {favoriteCount}."
                    : $"Znaleziono zapisane wersje: {savedVersions.Count}.";
        }
        catch {
            statusMessage = "Nie udało się odczytać listy zapisanych wersji.";
        }

        return items;
    }

    private void UpdateUploadVersionButton() {
        UploadVersionButton.IsEnabled = _viewModel.SelectedCalculator != null;
    }

    private void UploadVersionFromDisk() {
        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        var dialog = new OpenFileDialog {
            Title = "Wgraj wersję kalkulatora",
            Filter = "Pliki wersji (*.json)|*.json|Wszystkie pliki (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
            InitialDirectory = GetVersionImportInitialDirectory()
        };

        var dialogResult = dialog.ShowDialog(Window.GetWindow(this));
        if (dialogResult != true || string.IsNullOrWhiteSpace(dialog.FileName)) return;

        try {
            var importedVersionPath = ImportVersionFile(calculator, dialog.FileName);
            RefreshVersionListAndSelect(calculator, importedVersionPath);

            VersionStatusText.Text = $"Wgrano wersję: {Path.GetFileNameWithoutExtension(importedVersionPath)}.";
        }
        catch {
            MessageBox.Show(
                Window.GetWindow(this)!,
                "Nie udało się wgrać wybranej wersji. Sprawdź, czy to poprawny plik JSON pasujący do tego kalkulatora.",
                "Nie udało się wgrać wersji",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private static string GetVersionImportInitialDirectory() {
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        if (!string.IsNullOrWhiteSpace(userProfile)) {
            var downloadsDirectory = Path.Combine(userProfile, "Downloads");
            if (Directory.Exists(downloadsDirectory))
                return downloadsDirectory;
        }

        var desktopDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (!string.IsNullOrWhiteSpace(desktopDirectory) && Directory.Exists(desktopDirectory))
            return desktopDirectory;

        var documentsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        if (!string.IsNullOrWhiteSpace(documentsDirectory) && Directory.Exists(documentsDirectory))
            return documentsDirectory;

        return AppDomain.CurrentDomain.BaseDirectory;
    }

    private static string ImportVersionFile(CalculatorInfo calculator, string sourceFilePath) {
        if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
            throw new FileNotFoundException("Nie znaleziono wskazanego pliku wersji.");

        var version = CalculatorVersionService.Load(sourceFilePath);
        CalculatorVersionService.ValidateBelongsToCalculator(version, calculator.FilePath);

        var versionsDirectory = CalculatorVersionService.GetVersionsDirectory(calculator.FilePath);
        Directory.CreateDirectory(versionsDirectory);

        var destinationFilePath = GetUniqueImportedVersionPath(versionsDirectory, sourceFilePath);

        if (!string.Equals(Path.GetFullPath(sourceFilePath), Path.GetFullPath(destinationFilePath),
                StringComparison.OrdinalIgnoreCase))
            File.Copy(sourceFilePath, destinationFilePath, false);

        return destinationFilePath;
    }

    private static string GetUniqueImportedVersionPath(string versionsDirectory, string sourceFilePath) {
        var sourceFileName = Path.GetFileName(sourceFilePath);
        var safeFileName = string.IsNullOrWhiteSpace(sourceFileName)
            ? $"wersja_import_{DateTime.Now:yyyyMMdd_HHmmss}.json"
            : sourceFileName;

        var destinationFilePath = Path.Combine(versionsDirectory, safeFileName);
        var sourceFullPath = Path.GetFullPath(sourceFilePath);
        var destinationFullPath = Path.GetFullPath(destinationFilePath);

        if (string.Equals(sourceFullPath, destinationFullPath, StringComparison.OrdinalIgnoreCase))
            return destinationFilePath;

        if (!File.Exists(destinationFilePath))
            return destinationFilePath;

        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(safeFileName);
        var extension = Path.GetExtension(safeFileName);
        if (string.IsNullOrWhiteSpace(extension))
            extension = ".json";

        for (var index = 2; index < 1000; index++) {
            var candidatePath = Path.Combine(versionsDirectory, $"{fileNameWithoutExtension}_{index}{extension}");
            if (!File.Exists(candidatePath))
                return candidatePath;
        }

        return Path.Combine(versionsDirectory, $"{fileNameWithoutExtension}_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
    }

    private void RefreshVersionListAndSelect(CalculatorInfo calculator, string selectedVersionPath) {
        var versions = LoadVersionItems(calculator, out var versionStatusMessage);
        VersionList.ItemsSource = versions;
        VersionStatusText.Text = versionStatusMessage;

        if (!string.IsNullOrWhiteSpace(selectedVersionPath))
            foreach (var version in versions)
                if (!version.IsCurrentFile &&
                    string.Equals(version.FilePath, selectedVersionPath, StringComparison.OrdinalIgnoreCase)) {
                    VersionList.SelectedItem = version;
                    VersionList.ScrollIntoView(version);
                    return;
                }

        VersionList.SelectedIndex = versions.Count > 0 ? 0 : -1;
    }

    private void OpenCurrentCalculator() {
        if (_viewModel.SelectedCalculator == null) return;

        CalculatorStartupVersionSelection.Clear(_viewModel.SelectedCalculator.FilePath);

        if (_viewModel.OpenCalculatorCommand.CanExecute(null))
            _viewModel.OpenCalculatorCommand.Execute(null);
    }

    private void OpenSelectedVersion() {
        var calculator = _viewModel.SelectedCalculator;
        var selectedVersion = VersionList.SelectedItem as VersionListEntry;

        if (calculator == null || selectedVersion == null) return;

        if (selectedVersion.IsCurrentFile)
            CalculatorStartupVersionSelection.Clear(calculator.FilePath);
        else
            CalculatorStartupVersionSelection.Set(calculator.FilePath, selectedVersion.FilePath);

        if (_viewModel.OpenCalculatorCommand.CanExecute(null))
            _viewModel.OpenCalculatorCommand.Execute(null);
    }

    private static string? GetStringProperty(object source, string propertyName) {
        return source.GetType()
            .GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase)
            ?.GetValue(source) as string;
    }

    private static DateTime? GetDateTimeProperty(object source, string propertyName) {
        var value = source.GetType()
            .GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase)
            ?.GetValue(source);

        return value switch {
            DateTime dateTime => dateTime,
            DateTimeOffset dateTimeOffset => dateTimeOffset.LocalDateTime,
            _ => null
        };
    }


    internal static string FormatLastModified(string filePath) {
        try {
            return File.GetLastWriteTime(filePath).ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture);
        }
        catch {
            return string.Empty;
        }
    }

    internal static string FormatFileSize(string filePath) {
        try {
            var bytes = new FileInfo(filePath).Length;
            if (bytes < 1024)
                return $"{bytes} B";

            var kilobytes = bytes / 1024d;
            if (kilobytes < 1024)
                return $"{kilobytes:0.#} KB";

            return $"{kilobytes / 1024d:0.#} MB";
        }
        catch {
            return string.Empty;
        }
    }
}

public sealed class CalculatorListEntry {
    internal CalculatorListEntry(CalculatorInfo calculator) {
        Calculator = calculator;
    }

    internal CalculatorInfo Calculator { get; }
    internal string FilePath => Calculator.FilePath;

    public string DisplayName => Calculator.DisplayName;
    public string ModifiedText => CalculatorListView.FormatLastModified(Calculator.FilePath);
    public string SizeText => CalculatorListView.FormatFileSize(Calculator.FilePath);
}

public sealed class VersionListEntry {
    public required string DisplayName { get; init; }
    public required string Details { get; init; }
    public required string FilePath { get; init; }
    public required string SizeText { get; init; }
    public required string Icon { get; init; }
    public DateTime CreatedAt { get; init; }
    public bool IsCurrentFile { get; init; }
    public bool IsFavorite { get; init; }
    public Visibility ActionButtonsVisibility => IsCurrentFile ? Visibility.Collapsed : Visibility.Visible;
    public string FavoriteIcon => IsFavorite ? "\uE735" : "\uE734";
    public string FavoriteToolTip => IsFavorite ? "Usuń z ulubionych" : "Dodaj do ulubionych";

    public static VersionListEntry CreateCurrent(CalculatorInfo calculator) {
        return new VersionListEntry {
            DisplayName = "Aktualny plik kalkulatora",
            Details = CalculatorListView.FormatLastModified(calculator.FilePath),
            FilePath = calculator.FilePath,
            SizeText = CalculatorListView.FormatFileSize(calculator.FilePath),
            Icon = "\uE8A5",
            CreatedAt = File.Exists(calculator.FilePath)
                ? File.GetLastWriteTime(calculator.FilePath)
                : DateTime.MinValue,
            IsCurrentFile = true,
            IsFavorite = false
        };
    }

    public static VersionListEntry CreateSavedVersion(
        string filePath,
        string displayName,
        DateTime createdAt,
        bool isFavorite) {
        return new VersionListEntry {
            DisplayName = displayName,
            Details = isFavorite
                ? $"★ Ulubiona · Zapisano: {createdAt.ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture)}"
                : $"Zapisano: {createdAt.ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture)}",
            FilePath = filePath,
            SizeText = CalculatorListView.FormatFileSize(filePath),
            Icon = isFavorite ? "\uE735" : "\uE8AB",
            CreatedAt = createdAt,
            IsCurrentFile = false,
            IsFavorite = isFavorite
        };
    }
}

internal static class VersionFavoritesStore {
    private static readonly object SyncRoot = new();

    private static string StoreFilePath {
        get {
            var localApplicationData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var directory = string.IsNullOrWhiteSpace(localApplicationData)
                ? AppDomain.CurrentDomain.BaseDirectory
                : Path.Combine(localApplicationData, "CalculatorHost");

            return Path.Combine(directory, "version-favorites.txt");
        }
    }

    public static bool IsFavorite(string filePath) {
        if (string.IsNullOrWhiteSpace(filePath))
            return false;

        lock (SyncRoot) {
            return LoadFavorites().Contains(NormalizePath(filePath));
        }
    }

    public static void SetFavorite(string filePath, bool isFavorite) {
        if (string.IsNullOrWhiteSpace(filePath))
            return;

        lock (SyncRoot) {
            var favorites = LoadFavorites();
            var normalizedPath = NormalizePath(filePath);

            if (isFavorite)
                favorites.Add(normalizedPath);
            else
                favorites.Remove(normalizedPath);

            SaveFavorites(favorites);
        }
    }

    public static void RemoveFavorite(string filePath) {
        SetFavorite(filePath, false);
    }

    private static HashSet<string> LoadFavorites() {
        var favorites = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var storeFilePath = StoreFilePath;

        if (!File.Exists(storeFilePath))
            return favorites;

        try {
            foreach (var line in File.ReadAllLines(storeFilePath)) {
                var normalizedLine = NormalizePath(line);
                if (!string.IsNullOrWhiteSpace(normalizedLine))
                    favorites.Add(normalizedLine);
            }
        }
        catch {
            // Jeśli plik ulubionych jest chwilowo niedostępny, lista działa dalej bez ulubionych.
        }

        return favorites;
    }

    private static void SaveFavorites(HashSet<string> favorites) {
        var storeFilePath = StoreFilePath;
        var directory = Path.GetDirectoryName(storeFilePath);

        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);

        File.WriteAllLines(storeFilePath, favorites.OrderBy(path => path, StringComparer.OrdinalIgnoreCase));
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