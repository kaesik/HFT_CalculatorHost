using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using CalculatorHost.Models;
using CalculatorHost.Services;
using CalculatorHost.ViewModels;
using Microsoft.Win32;

namespace CalculatorHost.Views;

public partial class CalculatorListView {
    private readonly MainViewModel _viewModel;
    private CancellationTokenSource? _cacheUpdateCancellation;
    private List<CalculatorListEntry> _calculatorListItems = new();
    private VersionListEntry? _currentlyRenamedVersion;
    private bool _isRefreshingCalculatorList;
    private bool _isSubscribed;
    private bool _isUpdatingCache;

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
        _cacheUpdateCancellation?.Cancel();

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
        if (VersionList.SelectedItem is VersionListEntry { IsEditing: true }) {
            e.Handled = true;
            return;
        }

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

    private void CopyVersionButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if ((sender as FrameworkElement)?.Tag is not VersionListEntry version || version.IsCurrentFile)
            return;

        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        try {
            var copyPath = CopyVersionFile(version.FilePath);
            RefreshVersionListAndSelect(calculator, copyPath);

            if (VersionList.SelectedItem is VersionListEntry copiedVersion &&
                string.Equals(copiedVersion.FilePath, copyPath, StringComparison.OrdinalIgnoreCase)) {
                BeginInlineRename(
                    copiedVersion,
                    Path.GetFileNameWithoutExtension(copyPath),
                    "Skopiowano wersję. Wpisz nazwę kopii na liście. Enter zapisuje, Esc zostawia obecną nazwę.");
                return;
            }

            VersionStatusText.Text = $"Skopiowano wersję: {Path.GetFileNameWithoutExtension(copyPath)}.";
        }
        catch {
            MessageBox.Show(
                Window.GetWindow(this)!,
                "Nie udało się utworzyć kopii wybranej wersji.",
                "Nie udało się skopiować wersji",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void RenameVersionButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if ((sender as FrameworkElement)?.Tag is not VersionListEntry version || version.IsCurrentFile)
            return;

        if (_viewModel.SelectedCalculator == null)
            return;

        BeginInlineRename(version);
    }

    private void ConfirmInlineRenameButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if ((sender as FrameworkElement)?.Tag is VersionListEntry version)
            ConfirmInlineRename(version);
    }

    private void CancelInlineRenameButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if ((sender as FrameworkElement)?.Tag is VersionListEntry version)
            CancelInlineRename(version);
    }

    private void InlineRenameTextBox_KeyDown(object sender, KeyEventArgs e) {
        if ((sender as FrameworkElement)?.DataContext is not VersionListEntry version)
            return;

        switch (e.Key) {
            case Key.Enter:
                e.Handled = true;
                ConfirmInlineRename(version);
                break;
            case Key.Escape:
                e.Handled = true;
                CancelInlineRename(version);
                break;
        }
    }

    private void InlineRenameTextBox_Loaded(object sender, RoutedEventArgs e) {
        if (sender is not TextBox textBox || textBox.DataContext is not VersionListEntry { IsEditing: true })
            return;

        textBox.Focus();
        textBox.SelectAll();
    }

    private void BeginInlineRename(
        VersionListEntry version,
        string? initialName = null,
        string? statusMessage = null) {
        if (version.IsCurrentFile)
            return;

        if (_currentlyRenamedVersion != null && !ReferenceEquals(_currentlyRenamedVersion, version))
            _currentlyRenamedVersion.CancelRename();

        _currentlyRenamedVersion = version;
        version.BeginRename(string.IsNullOrWhiteSpace(initialName)
            ? Path.GetFileNameWithoutExtension(version.FilePath)
            : initialName.Trim());

        VersionList.SelectedItem = version;
        VersionList.ScrollIntoView(version);
        VersionStatusText.Text = string.IsNullOrWhiteSpace(statusMessage)
            ? "Zmień nazwę na liście. Enter zapisuje, Esc anuluje."
            : statusMessage;

        FocusInlineRenameTextBox(version);
    }

    private void ConfirmInlineRename(VersionListEntry version) {
        if (version.IsCurrentFile)
            return;

        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        var newName = version.EditedName.Trim();
        if (string.IsNullOrWhiteSpace(newName)) {
            VersionStatusText.Text = "Podaj nazwę wersji.";
            FocusInlineRenameTextBox(version);
            return;
        }

        try {
            var renamedPath = RenameVersionFile(version.FilePath, newName);

            if (string.Equals(version.FilePath, renamedPath, StringComparison.OrdinalIgnoreCase)) {
                version.CancelRename();
                if (ReferenceEquals(_currentlyRenamedVersion, version))
                    _currentlyRenamedVersion = null;

                VersionStatusText.Text = $"Nazwa wersji bez zmian: {Path.GetFileNameWithoutExtension(renamedPath)}.";
                return;
            }

            VersionFavoritesStore.MoveFavorite(version.FilePath, renamedPath);
            VersionDefaultsStore.MoveDefault(calculator.FilePath, version.FilePath, renamedPath);

            if (ReferenceEquals(_currentlyRenamedVersion, version))
                _currentlyRenamedVersion = null;

            RefreshVersionListAndSelect(calculator, renamedPath);
            VersionStatusText.Text = $"Zmieniono nazwę wersji: {Path.GetFileNameWithoutExtension(renamedPath)}.";
        }
        catch (Exception exception) {
            VersionStatusText.Text = FormatExceptionMessage("Nie udało się zmienić nazwy wybranej wersji", exception);
            FocusInlineRenameTextBox(version);
        }
    }

    private void CancelInlineRename(VersionListEntry version) {
        version.CancelRename();

        if (ReferenceEquals(_currentlyRenamedVersion, version))
            _currentlyRenamedVersion = null;

        VersionStatusText.Text = "Anulowano zmianę nazwy.";
    }

    private void FocusInlineRenameTextBox(VersionListEntry version) {
        VersionList.Dispatcher.BeginInvoke(new Action(() => {
            VersionList.UpdateLayout();

            if (VersionList.ItemContainerGenerator.ContainerFromItem(version) is not DependencyObject container)
                return;

            var textBox = FindVisualChild<TextBox>(container, "InlineRenameTextBox");
            if (textBox == null || !textBox.IsVisible)
                return;

            textBox.Focus();
            textBox.SelectAll();
        }), DispatcherPriority.Background);
    }

    private static T? FindVisualChild<T>(DependencyObject parent, string? childName = null)
        where T : FrameworkElement {
        var childrenCount = VisualTreeHelper.GetChildrenCount(parent);

        for (var index = 0; index < childrenCount; index++) {
            var child = VisualTreeHelper.GetChild(parent, index);

            if (child is T typedChild &&
                (string.IsNullOrWhiteSpace(childName) ||
                 string.Equals(typedChild.Name, childName, StringComparison.Ordinal)))
                return typedChild;

            var foundChild = FindVisualChild<T>(child, childName);
            if (foundChild != null)
                return foundChild;
        }

        return null;
    }

    private void DefaultVersionButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if ((sender as FrameworkElement)?.Tag is not VersionListEntry version)
            return;

        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        if (version.IsDefault) {
            VersionStatusText.Text = version.IsCurrentFile
                ? "Aktualny plik kalkulatora jest już domyślny."
                : $"Ta wersja jest już domyślna: {version.DisplayName}.";
            return;
        }

        var message = version.IsCurrentFile
            ? "Ustawić aktualny plik kalkulatora jako domyślny? Przy otwieraniu kalkulatora z listy nie będzie automatycznie wczytywana żadna zapisana wersja."
            : $"Ustawić wersję „{version.DisplayName}” jako domyślną dla kalkulatora „{calculator.DisplayName}”? Przy otwieraniu kalkulatora z listy ta wersja będzie wczytywana automatycznie.";

        var result = MessageBox.Show(
            Window.GetWindow(this)!,
            message,
            "Ustaw jako domyślną",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
            return;

        try {
            if (version.IsCurrentFile)
                VersionDefaultsStore.ClearDefault(calculator.FilePath);
            else
                VersionDefaultsStore.SetDefault(calculator.FilePath, version.FilePath);

            RefreshVersionListAndSelect(calculator, version.FilePath);
            VersionStatusText.Text = version.IsCurrentFile
                ? "Ustawiono aktualny plik kalkulatora jako domyślny."
                : $"Ustawiono domyślną wersję: {version.DisplayName}.";
        }
        catch {
            MessageBox.Show(
                Window.GetWindow(this)!,
                "Nie udało się ustawić domyślnej wersji.",
                "Nie udało się ustawić domyślnej wersji",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void DeleteVersionButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if ((sender as FrameworkElement)?.Tag is not VersionListEntry version || version.IsCurrentFile)
            return;

        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        var isDefaultVersion = VersionDefaultsStore.IsDefault(calculator.FilePath, version.FilePath);
        var confirmationMessage = isDefaultVersion
            ? $"Usunąć wersję „{version.DisplayName}”? To jest aktualnie domyślna wersja dla tego kalkulatora. Po usunięciu domyślny będzie aktualny plik kalkulatora."
            : $"Usunąć wersję „{version.DisplayName}”?";

        var result = MessageBox.Show(
            Window.GetWindow(this)!,
            confirmationMessage,
            "Usuń wersję",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
            return;

        try {
            VersionFavoritesStore.RemoveFavorite(version.FilePath);

            if (isDefaultVersion)
                VersionDefaultsStore.ClearDefault(calculator.FilePath);

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

    private void VersionActionButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
        if (e.ClickCount > 1)
            e.Handled = true;
    }

    private void VersionMoreActionsButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if (sender is not Button button || button.ContextMenu == null)
            return;

        button.ContextMenu.PlacementTarget = button;
        button.ContextMenu.IsOpen = true;
    }

    private void VersionPanelOptionsButton_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        if (sender is not Button button || button.ContextMenu == null)
            return;

        button.ContextMenu.PlacementTarget = button;
        button.ContextMenu.IsOpen = true;
    }

    private void ShowVersionsFolderMenuItem_Click(object sender, RoutedEventArgs e) {
        e.Handled = true;

        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        try {
            var versionsDirectory = CalculatorVersionService.GetVersionsDirectory(calculator.FilePath);
            Directory.CreateDirectory(versionsDirectory);

            Process.Start(new ProcessStartInfo {
                FileName = versionsDirectory,
                UseShellExecute = true
            });

            VersionStatusText.Text = "Otworzono folder z wersjami.";
        }
        catch (Exception exception) {
            MessageBox.Show(
                Window.GetWindow(this)!,
                FormatExceptionMessage("Nie udało się otworzyć folderu z wersjami", exception),
                "Nie udało się otworzyć folderu",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
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
        var defaultVersionPath = VersionDefaultsStore.GetDefault(calculator.FilePath);
        var currentFileIsDefault = string.IsNullOrWhiteSpace(defaultVersionPath);

        var items = new List<VersionListEntry> {
            VersionListEntry.CreateCurrent(calculator, currentFileIsDefault)
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
                    VersionFavoritesStore.IsFavorite(filePath),
                    VersionDefaultsStore.IsDefault(calculator.FilePath, filePath)));
            }

            items.AddRange(savedVersions
                .OrderByDescending(version => version.IsDefault)
                .ThenByDescending(version => version.IsFavorite)
                .ThenByDescending(version => version.CreatedAt)
                .ThenBy(version => version.DisplayName, StringComparer.CurrentCultureIgnoreCase));

            var favoriteCount = savedVersions.Count(version => version.IsFavorite);
            var defaultCount = savedVersions.Count(version => version.IsDefault);

            statusMessage = savedVersions.Count == 0
                ? "Brak zapisanych wersji pasujących do tego kalkulatora."
                : CreateVersionStatusMessage(savedVersions.Count, favoriteCount, defaultCount);
        }
        catch {
            statusMessage = "Nie udało się odczytać listy zapisanych wersji.";
        }

        return items;
    }

    private static string CreateVersionStatusMessage(int savedVersionsCount, int favoriteCount, int defaultCount) {
        var parts = new List<string> { $"Znaleziono zapisane wersje: {savedVersionsCount}." };

        if (defaultCount > 0)
            parts.Add("Domyślna: 1.");

        if (favoriteCount > 0)
            parts.Add($"Ulubione: {favoriteCount}.");

        return string.Join(" ", parts);
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

    private static string CopyVersionFile(string sourceFilePath) {
        if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
            throw new FileNotFoundException("Nie znaleziono pliku wersji do skopiowania.");

        var directory = Path.GetDirectoryName(sourceFilePath);
        if (string.IsNullOrWhiteSpace(directory))
            directory = AppDomain.CurrentDomain.BaseDirectory;

        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(sourceFilePath);
        if (string.IsNullOrWhiteSpace(fileNameWithoutExtension))
            fileNameWithoutExtension = "wersja";

        var copyFilePath = GetUniqueVersionCopyPath(directory, fileNameWithoutExtension);
        File.Copy(sourceFilePath, copyFilePath, false);
        return copyFilePath;
    }

    private static string RenameVersionFile(string sourceFilePath, string newDisplayName) {
        if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
            throw new FileNotFoundException("Nie znaleziono pliku wersji do zmiany nazwy.");

        var directory = Path.GetDirectoryName(sourceFilePath);
        if (string.IsNullOrWhiteSpace(directory))
            directory = AppDomain.CurrentDomain.BaseDirectory;

        var safeFileNameWithoutExtension = CreateSafeVersionFileNameWithoutExtension(newDisplayName);
        var destinationFilePath = Path.Combine(directory, $"{safeFileNameWithoutExtension}.json");

        if (string.Equals(
                Path.GetFullPath(sourceFilePath),
                Path.GetFullPath(destinationFilePath),
                StringComparison.OrdinalIgnoreCase))
            return sourceFilePath;

        if (File.Exists(destinationFilePath))
            throw new InvalidOperationException("Istnieje już wersja o takiej nazwie.");

        File.Move(sourceFilePath, destinationFilePath);
        return destinationFilePath;
    }

    private static string CreateSafeVersionFileNameWithoutExtension(string displayName) {
        var safeName = Path.GetFileName(displayName.Trim());

        if (safeName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            safeName = Path.GetFileNameWithoutExtension(safeName);

        foreach (var invalidCharacter in Path.GetInvalidFileNameChars())
            safeName = safeName.Replace(invalidCharacter, '_');

        safeName = string.Join(" ", safeName
                .Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
            .Trim()
            .Trim('.');

        if (string.IsNullOrWhiteSpace(safeName))
            throw new InvalidOperationException("Podana nazwa wersji jest nieprawidłowa.");

        return safeName;
    }

    private static string GetUniqueVersionCopyPath(string directory, string fileNameWithoutExtension) {
        var candidatePath = Path.Combine(directory,
            $"{fileNameWithoutExtension}_kopia_{DateTime.Now:yyyyMMdd_HHmmss}.json");

        if (!File.Exists(candidatePath))
            return candidatePath;

        for (var index = 2; index < 1000; index++) {
            candidatePath = Path.Combine(directory,
                $"{fileNameWithoutExtension}_kopia_{DateTime.Now:yyyyMMdd_HHmmss}_{index}.json");
            if (!File.Exists(candidatePath))
                return candidatePath;
        }

        return Path.Combine(directory, $"{fileNameWithoutExtension}_kopia_{Guid.NewGuid():N}.json");
    }

    private void RefreshVersionListAndSelect(CalculatorInfo calculator, string selectedVersionPath) {
        var versions = LoadVersionItems(calculator, out var versionStatusMessage);
        VersionList.ItemsSource = versions;
        VersionStatusText.Text = versionStatusMessage;

        if (!string.IsNullOrWhiteSpace(selectedVersionPath))
            foreach (var version in versions)
                if (string.Equals(version.FilePath, selectedVersionPath, StringComparison.OrdinalIgnoreCase)) {
                    VersionList.SelectedItem = version;
                    VersionList.ScrollIntoView(version);
                    return;
                }

        VersionList.SelectedIndex = versions.Count > 0 ? 0 : -1;
    }

    private void OpenCurrentCalculator() {
        if (_isUpdatingCache) {
            CacheUpdateDetailText.Text = "Aktualizacja cache jest w toku. Anuluj ją lub poczekaj na zakończenie.";
            return;
        }

        var calculator = _viewModel.SelectedCalculator;
        if (calculator == null) return;

        var defaultVersionPath = VersionDefaultsStore.GetDefault(calculator.FilePath);

        if (!string.IsNullOrWhiteSpace(defaultVersionPath))
            CalculatorStartupVersionSelection.Set(calculator.FilePath, defaultVersionPath);
        else
            CalculatorStartupVersionSelection.Clear(calculator.FilePath);

        if (_viewModel.OpenCalculatorCommand.CanExecute(null))
            _viewModel.OpenCalculatorCommand.Execute(null);
    }

    private void OpenSelectedVersion() {
        if (_isUpdatingCache) {
            CacheUpdateDetailText.Text = "Aktualizacja cache jest w toku. Anuluj ją lub poczekaj na zakończenie.";
            return;
        }

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


    private async void UpdateAllCachesButton_Click(object sender, RoutedEventArgs e) {
        if (_isUpdatingCache)
            return;

        var calculators = _viewModel.Calculators
            .Where(calculator => File.Exists(calculator.FilePath))
            .ToList();

        if (calculators.Count == 0) {
            MessageBox.Show(
                Window.GetWindow(this)!,
                "Nie znaleziono kalkulatorów do aktualizacji.",
                "Aktualizacja cache",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        _isUpdatingCache = true;
        _cacheUpdateCancellation = new CancellationTokenSource();

        UpdateAllCachesButton.IsEnabled = false;
        CancelCacheUpdateButton.IsEnabled = true;
        CacheUpdatePanel.Visibility = Visibility.Visible;
        CacheUpdateOverallProgressBar.Value = 0.0;
        CacheUpdateStageProgressBar.Value = 0.0;
        CacheUpdateOverallPercentText.Text = "0%";
        CacheUpdateStagePercentText.Text = "0%";
        CacheUpdateTitleText.Text = "Aktualizacja pamięci kalkulatorów";
        CacheUpdateCalculatorText.Text = $"Przygotowanie · 0 z {calculators.Count}";
        CacheUpdateStageText.Text = "Sprawdzanie cache";
        CacheUpdateDetailText.Text = "Sprawdzanie, które kalkulatory wymagają aktualizacji…";
        CacheUpdateStatsText.Text = "Zaktualizowano: 0 · Aktualne: 0 · Błędy: 0";

        var cacheUpdateService = new CalculatorCacheUpdateService(
            new SheetLayoutCacheService(),
            App.ExcelWorker);

        var progress = new Progress<CalculatorCacheUpdateProgress>(UpdateCacheProgressUi);

        try {
            var result = await cacheUpdateService.UpdateAllAsync(
                calculators,
                progress,
                _cacheUpdateCancellation.Token);

            CacheUpdateOverallProgressBar.Value = result.WasCancelled
                ? CacheUpdateOverallProgressBar.Value
                : 100.0;
            CacheUpdateOverallPercentText.Text = result.WasCancelled
                ? $"{CacheUpdateOverallProgressBar.Value:0}%"
                : "100%";
            CacheUpdateStageProgressBar.Value = 100.0;
            CacheUpdateStagePercentText.Text = "100%";
            CacheUpdateCalculatorText.Text = result.WasCancelled
                ? "Aktualizacja anulowana"
                : "Aktualizacja zakończona";
            CacheUpdateStageText.Text = result.WasCancelled
                ? "Anulowano"
                : "Gotowe";
            CacheUpdateDetailText.Text = result.WasCancelled
                ? "Przerwano po zakończeniu bieżącej operacji."
                : "Cache wszystkich wymagających aktualizacji kalkulatorów jest gotowy.";
            CacheUpdateStatsText.Text =
                $"Zaktualizowano: {result.UpdatedCount} · " +
                $"Aktualne: {result.SkippedCount} · " +
                $"Błędy: {result.FailedCount}";
        }
        catch (Exception exception) {
            CacheUpdateStageText.Text = "Błąd aktualizacji";
            CacheUpdateDetailText.Text = FormatExceptionMessage(
                "Nie udało się zakończyć aktualizacji cache",
                exception);
        }
        finally {
            _isUpdatingCache = false;
            UpdateAllCachesButton.IsEnabled = true;
            CancelCacheUpdateButton.IsEnabled = false;
            _cacheUpdateCancellation?.Dispose();
            _cacheUpdateCancellation = null;
        }
    }

    private void CancelCacheUpdateButton_Click(object sender, RoutedEventArgs e) {
        if (!_isUpdatingCache || _cacheUpdateCancellation == null)
            return;

        CancelCacheUpdateButton.IsEnabled = false;
        CacheUpdateStageText.Text = "Anulowanie…";
        CacheUpdateDetailText.Text =
            "Aktualizacja zostanie zatrzymana po zakończeniu bieżącej operacji na pliku.";
        _cacheUpdateCancellation.Cancel();
    }

    private void CloseCacheUpdatePanelButton_Click(object sender, RoutedEventArgs e) {
        if (_isUpdatingCache)
            return;

        CacheUpdatePanel.Visibility = Visibility.Collapsed;
    }

    private void UpdateCacheProgressUi(CalculatorCacheUpdateProgress progress) {
        CacheUpdateOverallProgressBar.Value = progress.OverallPercentage;
        CacheUpdateStageProgressBar.Value = progress.StagePercentage;
        CacheUpdateOverallPercentText.Text = $"{progress.OverallPercentage:0}%";
        CacheUpdateStagePercentText.Text = $"{progress.StagePercentage:0}%";
        CacheUpdateCalculatorText.Text =
            $"{progress.CalculatorName} · {progress.CurrentIndex} z {progress.TotalCount}";
        CacheUpdateStageText.Text = progress.Stage;
        CacheUpdateDetailText.Text = progress.Detail;
        CacheUpdateStatsText.Text =
            $"Zaktualizowano: {progress.UpdatedCount} · " +
            $"Aktualne: {progress.SkippedCount} · " +
            $"Błędy: {progress.FailedCount}";
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


    private static string FormatExceptionMessage(string prefix, Exception exception) {
        var messages = new List<string>();
        var currentException = exception;

        while (currentException != null) {
            if (!string.IsNullOrWhiteSpace(currentException.Message))
                messages.Add(currentException.Message);

            currentException = currentException.InnerException;
        }

        return messages.Count == 0
            ? prefix
            : $"{prefix}: {string.Join(Environment.NewLine, messages)}";
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

public sealed class VersionListEntry : INotifyPropertyChanged {
    private string _editedName = string.Empty;
    private bool _isEditing;

    public required string DisplayName { get; init; }
    public required string Details { get; init; }
    public required string FilePath { get; init; }
    public required string SizeText { get; init; }
    public required string Icon { get; init; }
    public DateTime CreatedAt { get; init; }
    public bool IsCurrentFile { get; init; }
    public bool IsFavorite { get; init; }
    public bool IsDefault { get; init; }

    public string EditedName {
        get => _editedName;
        set {
            if (string.Equals(_editedName, value, StringComparison.Ordinal))
                return;

            _editedName = value;
            OnPropertyChanged();
        }
    }

    public bool IsEditing {
        get => _isEditing;
        private set {
            if (_isEditing == value)
                return;

            _isEditing = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(DisplayNameVisibility));
            OnPropertyChanged(nameof(RenameEditorVisibility));
            OnPropertyChanged(nameof(ActionButtonsVisibility));
        }
    }

    public Visibility DisplayNameVisibility => IsEditing ? Visibility.Collapsed : Visibility.Visible;
    public Visibility RenameEditorVisibility => IsEditing ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ActionButtonsVisibility => IsEditing ? Visibility.Collapsed : Visibility.Visible;
    public Visibility SavedVersionButtonVisibility => IsCurrentFile ? Visibility.Collapsed : Visibility.Visible;
    public Visibility CurrentFileDefaultButtonVisibility => IsCurrentFile ? Visibility.Visible : Visibility.Collapsed;
    public Visibility MoreActionsButtonVisibility => IsCurrentFile ? Visibility.Collapsed : Visibility.Visible;

    public Visibility DefaultButtonVisibility =>
        !IsCurrentFile && !IsDefault ? Visibility.Visible : Visibility.Collapsed;

    public string FavoriteIcon => IsFavorite ? "\uE735" : "\uE734";
    public string FavoriteToolTip => IsFavorite ? "Usuń z ulubionych" : "Dodaj do ulubionych";
    public string DefaultIcon => IsDefault ? "\uE840" : "\uE718";
    public string DefaultToolTip => IsDefault ? "Domyślna wersja - przypięta" : "Ustaw jako domyślną";

    public event PropertyChangedEventHandler? PropertyChanged;

    public void BeginRename(string currentName) {
        if (IsCurrentFile)
            return;

        EditedName = currentName;
        IsEditing = true;
    }

    public void CancelRename() {
        IsEditing = false;
        EditedName = string.Empty;
    }

    public static VersionListEntry CreateCurrent(CalculatorInfo calculator, bool isDefault) {
        var lastModifiedText = CalculatorListView.FormatLastModified(calculator.FilePath);

        return new VersionListEntry {
            DisplayName = "Aktualny plik kalkulatora",
            Details = isDefault
                ? $"✓ Domyślny · {lastModifiedText}"
                : lastModifiedText,
            FilePath = calculator.FilePath,
            SizeText = CalculatorListView.FormatFileSize(calculator.FilePath),
            Icon = "\uE8A5",
            CreatedAt = File.Exists(calculator.FilePath)
                ? File.GetLastWriteTime(calculator.FilePath)
                : DateTime.MinValue,
            IsCurrentFile = true,
            IsFavorite = false,
            IsDefault = isDefault
        };
    }

    public static VersionListEntry CreateSavedVersion(
        string filePath,
        string displayName,
        DateTime createdAt,
        bool isFavorite,
        bool isDefault) {
        return new VersionListEntry {
            DisplayName = displayName,
            Details = CreateSavedVersionDetails(createdAt, isFavorite, isDefault),
            FilePath = filePath,
            SizeText = CalculatorListView.FormatFileSize(filePath),
            Icon = isDefault ? "\uE840" : isFavorite ? "\uE735" : "\uE8AB",
            CreatedAt = createdAt,
            IsCurrentFile = false,
            IsFavorite = isFavorite,
            IsDefault = isDefault
        };
    }

    private static string CreateSavedVersionDetails(DateTime createdAt, bool isFavorite, bool isDefault) {
        var parts = new List<string>();

        if (isDefault)
            parts.Add("✓ Domyślna");

        if (isFavorite)
            parts.Add("★ Ulubiona");

        parts.Add($"Zapisano: {createdAt.ToString("dd.MM.yyyy HH:mm", CultureInfo.CurrentCulture)}");

        return string.Join(" · ", parts);
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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

    public static void MoveFavorite(string oldFilePath, string newFilePath) {
        if (string.IsNullOrWhiteSpace(oldFilePath) || string.IsNullOrWhiteSpace(newFilePath))
            return;

        lock (SyncRoot) {
            var favorites = LoadFavorites();
            var oldNormalizedPath = NormalizePath(oldFilePath);

            if (favorites.Remove(oldNormalizedPath))
                favorites.Add(NormalizePath(newFilePath));

            SaveFavorites(favorites);
        }
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

internal static class VersionDefaultsStore {
    private static readonly object SyncRoot = new();

    private static string StoreFilePath {
        get {
            var localApplicationData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var directory = string.IsNullOrWhiteSpace(localApplicationData)
                ? AppDomain.CurrentDomain.BaseDirectory
                : Path.Combine(localApplicationData, "CalculatorHost");

            return Path.Combine(directory, "version-defaults.tsv");
        }
    }

    public static string? GetDefault(string calculatorFilePath) {
        if (string.IsNullOrWhiteSpace(calculatorFilePath))
            return null;

        lock (SyncRoot) {
            var defaults = LoadDefaults();
            var calculatorKey = NormalizePath(calculatorFilePath);

            if (!defaults.TryGetValue(calculatorKey, out var versionFilePath) ||
                string.IsNullOrWhiteSpace(versionFilePath))
                return null;

            if (File.Exists(versionFilePath))
                return versionFilePath;

            defaults.Remove(calculatorKey);
            SaveDefaults(defaults);
            return null;
        }
    }

    public static bool IsDefault(string calculatorFilePath, string versionFilePath) {
        if (string.IsNullOrWhiteSpace(calculatorFilePath) || string.IsNullOrWhiteSpace(versionFilePath))
            return false;

        var currentDefault = GetDefault(calculatorFilePath);
        return !string.IsNullOrWhiteSpace(currentDefault) &&
               string.Equals(
                   NormalizePath(currentDefault),
                   NormalizePath(versionFilePath),
                   StringComparison.OrdinalIgnoreCase);
    }

    public static void SetDefault(string calculatorFilePath, string versionFilePath) {
        if (string.IsNullOrWhiteSpace(calculatorFilePath) || string.IsNullOrWhiteSpace(versionFilePath))
            return;

        lock (SyncRoot) {
            var defaults = LoadDefaults();
            defaults[NormalizePath(calculatorFilePath)] = NormalizePath(versionFilePath);
            SaveDefaults(defaults);
        }
    }

    public static void ClearDefault(string calculatorFilePath) {
        if (string.IsNullOrWhiteSpace(calculatorFilePath))
            return;

        lock (SyncRoot) {
            var defaults = LoadDefaults();
            defaults.Remove(NormalizePath(calculatorFilePath));
            SaveDefaults(defaults);
        }
    }

    public static void MoveDefault(string calculatorFilePath, string oldVersionFilePath, string newVersionFilePath) {
        if (string.IsNullOrWhiteSpace(calculatorFilePath) ||
            string.IsNullOrWhiteSpace(oldVersionFilePath) ||
            string.IsNullOrWhiteSpace(newVersionFilePath))
            return;

        lock (SyncRoot) {
            var defaults = LoadDefaults();
            var calculatorKey = NormalizePath(calculatorFilePath);

            if (!defaults.TryGetValue(calculatorKey, out var currentDefaultPath))
                return;

            if (!string.Equals(
                    NormalizePath(currentDefaultPath),
                    NormalizePath(oldVersionFilePath),
                    StringComparison.OrdinalIgnoreCase))
                return;

            defaults[calculatorKey] = NormalizePath(newVersionFilePath);
            SaveDefaults(defaults);
        }
    }

    private static Dictionary<string, string> LoadDefaults() {
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var storeFilePath = StoreFilePath;

        if (!File.Exists(storeFilePath))
            return defaults;

        try {
            foreach (var line in File.ReadAllLines(storeFilePath)) {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split('	', 2);
                if (parts.Length != 2)
                    continue;

                var calculatorPath = NormalizePath(parts[0]);
                var versionPath = NormalizePath(parts[1]);

                if (!string.IsNullOrWhiteSpace(calculatorPath) && !string.IsNullOrWhiteSpace(versionPath))
                    defaults[calculatorPath] = versionPath;
            }
        }
        catch {
            // Jeśli plik domyślnych wersji jest chwilowo niedostępny, lista działa dalej bez domyślnej wersji.
        }

        return defaults;
    }

    private static void SaveDefaults(Dictionary<string, string> defaults) {
        var storeFilePath = StoreFilePath;
        var directory = Path.GetDirectoryName(storeFilePath);

        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);

        var lines = defaults
            .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .Select(pair => $"{pair.Key}	{pair.Value}");

        File.WriteAllLines(storeFilePath, lines);
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