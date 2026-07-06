using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using CalculatorHost.Models;
using CalculatorHost.Services;

namespace CalculatorHost.ViewModels;

public static class CalculatorStartupVersionSelection {
    private static readonly object SyncRoot = new();
    private static readonly Dictionary<string, string> SelectedVersionPaths = new(StringComparer.OrdinalIgnoreCase);

    public static void Set(string calculatorFilePath, string versionFilePath) {
        if (string.IsNullOrWhiteSpace(calculatorFilePath) || string.IsNullOrWhiteSpace(versionFilePath)) return;

        lock (SyncRoot) {
            SelectedVersionPaths[calculatorFilePath] = versionFilePath;
        }
    }

    public static void Clear(string calculatorFilePath) {
        if (string.IsNullOrWhiteSpace(calculatorFilePath)) return;

        lock (SyncRoot) {
            SelectedVersionPaths.Remove(calculatorFilePath);
        }
    }

    public static string? Take(string calculatorFilePath) {
        if (string.IsNullOrWhiteSpace(calculatorFilePath)) return null;

        lock (SyncRoot) {
            if (!SelectedVersionPaths.Remove(calculatorFilePath, out var versionFilePath))
                return null;

            return versionFilePath;
        }
    }
}

public class CalculatorViewModel : INotifyPropertyChanged, IDisposable {
    private readonly Dictionary<(int Row, int Column), string> _committedCellValues = [];
    private readonly ExcelSessionService _excelSession;
    private readonly Dictionary<(int Row, int Column), string> _pendingCellValues = [];
    private readonly SheetLayoutCacheService _sheetLayoutCache;
    private readonly SheetReaderService _sheetReader;
    private readonly ExcelWorker _worker;
    private readonly WorkingCopyService _workingCopy;
    private CalculatorInfo? _calculatorInfo;
    private string _calculatorName = string.Empty;
    private bool _disposed;
    private string _errorMessage = string.Empty;
    private bool _hasPendingChanges;
    private bool _isLoading;
    private ObservableCollection<MacroButtonConfig> _macroButtons = [];
    private string _operationPerformanceMessage = string.Empty;
    private string _performanceMessage = string.Empty;
    private SheetModel? _sheetModel;
    private string _statusMessage = string.Empty;

    public CalculatorViewModel(
        ExcelSessionService excelSession,
        SheetReaderService sheetReader,
        WorkingCopyService workingCopy,
        SheetLayoutCacheService sheetLayoutCache,
        ExcelWorker worker) {
        _excelSession = excelSession;
        _sheetReader = sheetReader;
        _workingCopy = workingCopy;
        _sheetLayoutCache = sheetLayoutCache;
        _worker = worker;

        CalculateCommand = new AsyncRelayCommand(CalculateAsync, () => !IsLoading && !HasError);
        RunMacroCommand = new AsyncRelayCommand<MacroButtonConfig>(RunMacroAsync, _ => !IsLoading && !HasError);
    }

    public SheetModel? SheetModel {
        get => _sheetModel;
        private set {
            _sheetModel = value;
            OnPropertyChanged();
        }
    }

    public ObservableCollection<MacroButtonConfig> MacroButtons {
        get => _macroButtons;
        private set {
            _macroButtons = value;
            OnPropertyChanged();
        }
    }

    public bool IsLoading {
        get => _isLoading;
        private set {
            if (_isLoading == value) return;

            _isLoading = value;
            OnPropertyChanged();
            CommandManager.InvalidateRequerySuggested();
        }
    }

    public string ErrorMessage {
        get => _errorMessage;
        private set {
            if (_errorMessage == value) return;

            _errorMessage = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasError));
            CommandManager.InvalidateRequerySuggested();
        }
    }

    public bool HasError => !string.IsNullOrWhiteSpace(ErrorMessage);

    public bool HasPendingChanges {
        get => _hasPendingChanges;
        private set {
            if (_hasPendingChanges == value) return;

            _hasPendingChanges = value;
            OnPropertyChanged();
        }
    }

    public string CalculatorName {
        get => _calculatorName;
        private set {
            _calculatorName = value;
            OnPropertyChanged();
        }
    }

    public string StatusMessage {
        get => _statusMessage;
        private set {
            if (_statusMessage == value) return;

            _statusMessage = value;
            OnPropertyChanged();
        }
    }

    public string PerformanceMessage {
        get => _performanceMessage;
        private set {
            if (_performanceMessage == value) return;

            _performanceMessage = value;
            OnPropertyChanged();
        }
    }

    public AsyncRelayCommand CalculateCommand { get; }
    public AsyncRelayCommand<MacroButtonConfig> RunMacroCommand { get; }

    public void Dispose() {
        if (_disposed) return;

        _disposed = true;

        try {
            _worker.InvokeAsync(() => {
                _excelSession.CloseWorkbook();
                _excelSession.Dispose();
            }).Wait(TimeSpan.FromSeconds(15));
        }
        catch {
            // ignored
        }

        _workingCopy.CleanCurrentSession();
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action? CloseRequested;

    public async Task LoadCalculatorAsync(CalculatorInfo calculatorInfo) {
        if (_disposed) return;

        _calculatorInfo = calculatorInfo;
        CalculatorName = calculatorInfo.DisplayName;
        IsLoading = true;
        ErrorMessage = string.Empty;
        StatusMessage = "Tworzenie kopii roboczej…";
        ClearPerformanceMessage();

        var operationName = "tworzenia kopii roboczej";

        try {
            if (!File.Exists(calculatorInfo.FilePath))
                throw new FileNotFoundException("Nie znaleziono wskazanego pliku kalkulatora.",
                    calculatorInfo.FilePath);

            var workingCopyStopwatch = Stopwatch.StartNew();
            var workingPath = _workingCopy.CreateWorkingCopy(calculatorInfo.FilePath);
            workingCopyStopwatch.Stop();

            operationName = "uruchamiania programu Excel i otwierania skoroszytu";
            StatusMessage = "Uruchamianie Excela…";
            var openingStopwatch = Stopwatch.StartNew();
            await _worker.InvokeAsync(() => _excelSession.OpenSession(workingPath));
            openingStopwatch.Stop();

            operationName = "sprawdzania pamięci układu arkusza";
            StatusMessage = "Sprawdzanie pamięci układu…";
            var cacheLoadStopwatch = Stopwatch.StartNew();
            var isLayoutLoadedFromCache = _sheetLayoutCache.TryLoad(calculatorInfo, out var cachedModel);
            cacheLoadStopwatch.Stop();

            operationName = isLayoutLoadedFromCache
                ? "odświeżania wartości arkusza"
                : "odczytu pierwszego arkusza";
            StatusMessage = isLayoutLoadedFromCache
                ? "Odczyt wartości arkusza…"
                : "Odczyt arkusza…";

            var readingStopwatch = Stopwatch.StartNew();
            SheetModel model;

            if (isLayoutLoadedFromCache && cachedModel != null)
                model = await _worker.InvokeAsync(() => _sheetReader.RefreshCellValues(_excelSession, cachedModel));
            else
                model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));

            readingStopwatch.Stop();

            var cacheSaveMessage = string.Empty;

            if (!isLayoutLoadedFromCache) {
                var cacheSaveStopwatch = Stopwatch.StartNew();
                var isCacheSaved = _sheetLayoutCache.TrySave(calculatorInfo, model);
                cacheSaveStopwatch.Stop();

                cacheSaveMessage = isCacheSaved
                    ? $" · Zapis cache: {FormatDuration(cacheSaveStopwatch.Elapsed)}"
                    : " · Zapis cache: niepowodzenie";
            }

            operationName = "wczytywania konfiguracji makr";
            var macroConfigurationStopwatch = Stopwatch.StartNew();
            var buttons = MacroConfigService.LoadForCalculator(calculatorInfo.FilePath);
            macroConfigurationStopwatch.Stop();

            if (_disposed) return;

            var readMessage = isLayoutLoadedFromCache
                ? $"Cache + wartości: {FormatDuration(readingStopwatch.Elapsed)}"
                : $"Pełny odczyt arkusza: {FormatDuration(readingStopwatch.Elapsed)}";

            var cacheMessage = isLayoutLoadedFromCache
                ? $"Cache: użyty ({FormatDuration(cacheLoadStopwatch.Elapsed)})"
                : $"Cache: brak ({FormatDuration(cacheLoadStopwatch.Elapsed)})";

            UpdateCommittedCellValues(model);
            ClearPendingChanges();
            SetOperationPerformanceMessage(
                $"Kopia: {FormatDuration(workingCopyStopwatch.Elapsed)} · " +
                $"Excel: {FormatDuration(openingStopwatch.Elapsed)} · " +
                $"{cacheMessage} · " +
                $"{readMessage}" +
                $"{cacheSaveMessage} · " +
                $"Konfiguracja makr: {FormatDuration(macroConfigurationStopwatch.Elapsed)}");
            SheetModel = model;
            MacroButtons = new ObservableCollection<MacroButtonConfig>(buttons);

            var startupVersionPath = CalculatorStartupVersionSelection.Take(calculatorInfo.FilePath);

            if (!string.IsNullOrWhiteSpace(startupVersionPath)) {
                operationName = "wczytywania wybranej wersji";
                StatusMessage = "Wczytywanie wybranej wersji…";
                StatusMessage = await ApplyVersionFileAsync(startupVersionPath);
            }
            else
                StatusMessage = string.Empty;
        }
        catch (Exception exception) {
            if (!_disposed) {
                ErrorMessage = FormatExceptionMessage($"Błąd podczas {operationName}", exception);
                StatusMessage = string.Empty;
            }
        }
        finally {
            if (!_disposed)
                IsLoading = false;
        }
    }

    public void SetPendingCellValue(int row, int column, string value) {
        if (_disposed || IsLoading || HasError || SheetModel == null) return;

        var originalValue = _committedCellValues.TryGetValue((row, column), out var committedValue)
            ? committedValue
            : SheetModel.Cells
                .FirstOrDefault(cell => cell.Row == row && cell.Column == column)
                ?.DisplayText ?? string.Empty;

        if (string.Equals(NormalizeDropdownText(originalValue), NormalizeDropdownText(value),
                StringComparison.OrdinalIgnoreCase))
            _pendingCellValues.Remove((row, column));
        else
            _pendingCellValues[(row, column)] = value;

        HasPendingChanges = _pendingCellValues.Count > 0;
        StatusMessage = HasPendingChanges
            ? "Wprowadzono zmiany — kliknij „Przelicz”."
            : string.Empty;
    }

    public string GetDefaultVersionFilePath() {
        if (_calculatorInfo == null)
            throw new InvalidOperationException("Brak informacji o pliku kalkulatora.");

        return CalculatorVersionService.CreateDefaultVersionFilePath(_calculatorInfo.FilePath);
    }

    public string GetVersionsDirectory() {
        if (_calculatorInfo == null)
            throw new InvalidOperationException("Brak informacji o pliku kalkulatora.");

        return CalculatorVersionService.GetVersionsDirectory(_calculatorInfo.FilePath);
    }

    public List<CalculatorVersionFileInfo> GetMatchingVersionFiles() {
        if (_calculatorInfo == null)
            throw new InvalidOperationException("Brak informacji o pliku kalkulatora.");

        return CalculatorVersionService.FindMatchingVersionFiles(_calculatorInfo.FilePath);
    }

    public async Task SaveVersionAsync(string versionFilePath) {
        if (_disposed || IsLoading) return;

        try {
            ErrorMessage = string.Empty;

            if (SheetModel == null)
                throw new InvalidOperationException("Nie ma wczytanego arkusza, więc nie można zapisać wersji.");

            if (_calculatorInfo == null)
                throw new InvalidOperationException("Brak informacji o pliku kalkulatora.");

            var version = CreateVersionModel();
            CalculatorVersionService.Save(versionFilePath, version);
            StatusMessage = $"Zapisano wersję: {Path.GetFileName(versionFilePath)}";
        }
        catch (Exception exception) {
            if (!_disposed)
                StatusMessage = FormatExceptionMessage("Błąd zapisu wersji", exception);
        }

        await Task.CompletedTask;
    }

    public async Task LoadVersionAsync(string versionFilePath) {
        if (_disposed || IsLoading || HasError) return;

        IsLoading = true;
        ErrorMessage = string.Empty;
        StatusMessage = "Wczytywanie wersji…";
        ClearPerformanceMessage();

        try {
            StatusMessage = await ApplyVersionFileAsync(versionFilePath);
        }
        catch (Exception exception) {
            if (!_disposed)
                StatusMessage = FormatExceptionMessage("Błąd wczytywania wersji", exception);
        }
        finally {
            if (!_disposed)
                IsLoading = false;
        }
    }

    private async Task<string> ApplyVersionFileAsync(string versionFilePath) {
        if (SheetModel == null)
            throw new InvalidOperationException("Nie ma wczytanego arkusza, więc nie można wczytać wersji.");

        if (_calculatorInfo == null)
            throw new InvalidOperationException("Brak informacji o pliku kalkulatora.");

        var version = CalculatorVersionService.Load(versionFilePath);
        ValidateVersion(version);

        var currentInputCells = SheetModel.Cells
            .Where(cell => cell is { IsInput: true, IsMergedSlave: false })
            .ToDictionary(cell => (cell.Row, cell.Column), cell => cell);

        var values = version.Values
            .Where(value => value.Row > 0 && value.Column > 0)
            .Where(value => currentInputCells.ContainsKey((value.Row, value.Column)))
            .Select(value => new KeyValuePair<(int Row, int Column), string>(
                (value.Row, value.Column),
                value.Value))
            .ToList();

        var skippedValuesCount = version.Values.Count - values.Count;

        if (values.Count == 0)
            throw new InvalidOperationException(
                "Plik wersji nie zawiera pól, które istnieją jako edytowalne pola w aktualnym arkuszu.");

        var applyingStopwatch = Stopwatch.StartNew();
        await _worker.InvokeAsync(() => {
            WritePendingValues(values);
            _excelSession.Recalculate();
        });
        applyingStopwatch.Stop();

        var currentModel = SheetModel;

        if (currentModel == null)
            return string.Empty;

        var refreshStopwatch = Stopwatch.StartNew();
        var model = await _worker.InvokeAsync(() => _sheetReader.RefreshCellValues(_excelSession, currentModel));
        refreshStopwatch.Stop();

        if (_disposed)
            return string.Empty;

        UpdateCommittedCellValues(model);
        ClearPendingChanges();
        SetOperationPerformanceMessage(
            $"Wczytanie wersji: {FormatDuration(applyingStopwatch.Elapsed)} · " +
            $"Odświeżenie wartości: {FormatDuration(refreshStopwatch.Elapsed)}");
        SheetModel = model;

        return skippedValuesCount > 0
            ? $"Wczytano wersję: {Path.GetFileName(versionFilePath)} (zastosowano {values.Count} z {version.Values.Count} pól)"
            : $"Wczytano wersję: {Path.GetFileName(versionFilePath)}";
    }

    public void ShowExternalError(string prefix, Exception exception) {
        if (_disposed) return;

        ErrorMessage = FormatExceptionMessage(prefix, exception);
        StatusMessage = string.Empty;
        IsLoading = false;
    }

    public void RequestClose() {
        CloseRequested?.Invoke();
    }

    public void ReportRenderingDuration(TimeSpan duration) {
        if (_disposed) return;

        var renderingMessage = $"Renderowanie: {FormatDuration(duration)}";
        PerformanceMessage = string.IsNullOrWhiteSpace(_operationPerformanceMessage)
            ? renderingMessage
            : $"{_operationPerformanceMessage} · {renderingMessage}";
    }

    private async Task CalculateAsync() {
        if (_disposed || IsLoading || HasError) return;

        IsLoading = true;
        ErrorMessage = string.Empty;
        StatusMessage = "Przeliczanie…";
        ClearPerformanceMessage();

        var pendingValues = BuildPendingValuesWithSynchronizedDropdowns(_pendingCellValues.ToList());

        try {
            var recalculationStopwatch = Stopwatch.StartNew();
            await _worker.InvokeAsync(() => {
                _excelSession.ExecuteWithEventsEnabled(() => {
                    WritePendingValues(pendingValues);
                    _excelSession.Recalculate();
                    WritePendingValues(pendingValues);
                    _excelSession.Recalculate();
                });
            });
            recalculationStopwatch.Stop();

            var currentModel = SheetModel;

            if (currentModel == null)
                return;

            var refreshStopwatch = Stopwatch.StartNew();
            var model = await _worker.InvokeAsync(() => _sheetReader.RefreshCellValues(_excelSession, currentModel));
            refreshStopwatch.Stop();

            if (_disposed) return;

            UpdateCommittedCellValues(model);
            ClearPendingChanges();
            SetOperationPerformanceMessage(
                $"Przeliczenie: {FormatDuration(recalculationStopwatch.Elapsed)} · " +
                $"Odświeżenie wartości: {FormatDuration(refreshStopwatch.Elapsed)}");
            SheetModel = model;
            StatusMessage = string.Empty;
        }
        catch (Exception exception) {
            if (!_disposed) {
                ErrorMessage = FormatExceptionMessage("Błąd przeliczania arkusza", exception);
                StatusMessage = string.Empty;
            }
        }
        finally {
            if (!_disposed)
                IsLoading = false;
        }
    }

    private async Task RunMacroAsync(MacroButtonConfig? config) {
        if (config == null || _disposed || IsLoading || HasError) return;

        IsLoading = true;
        ErrorMessage = string.Empty;
        StatusMessage = $"Wykonywanie: {config.Label}…";
        ClearPerformanceMessage();

        var pendingValues = BuildPendingValuesWithSynchronizedDropdowns(_pendingCellValues.ToList());

        try {
            var macroStopwatch = Stopwatch.StartNew();
            await _worker.InvokeAsync(() => {
                _excelSession.ExecuteWithEventsEnabled(() => {
                    WritePendingValues(pendingValues);
                    _excelSession.RunMacroButton(config);
                    WritePendingValues(pendingValues);
                    _excelSession.Recalculate();
                });
            });
            macroStopwatch.Stop();

            var currentModel = SheetModel;
            var refreshStopwatch = Stopwatch.StartNew();
            SheetModel model;
            string refreshOperationName;

            if (config.RefreshLayoutAfterRun || currentModel == null) {
                model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));
                refreshOperationName = "Pełny odczyt arkusza";
            }
            else {
                model = await _worker.InvokeAsync(() => _sheetReader.RefreshCellValues(_excelSession, currentModel));
                refreshOperationName = "Odświeżenie wartości";
            }

            refreshStopwatch.Stop();

            if (_disposed) return;

            UpdateCommittedCellValues(model);
            ClearPendingChanges();
            SetOperationPerformanceMessage(
                $"Makro: {FormatDuration(macroStopwatch.Elapsed)} · " +
                $"{refreshOperationName}: {FormatDuration(refreshStopwatch.Elapsed)}");
            SheetModel = model;
            StatusMessage = string.Empty;
        }
        catch (Exception exception) {
            if (!_disposed) {
                ErrorMessage = FormatExceptionMessage($"Makro '{config.Label}' nie powiodło się", exception);
                StatusMessage = string.Empty;
            }
        }
        finally {
            if (!_disposed)
                IsLoading = false;
        }
    }

    private CalculatorVersionModel CreateVersionModel() {
        if (SheetModel == null)
            throw new InvalidOperationException("Nie ma wczytanego arkusza.");

        if (_calculatorInfo == null)
            throw new InvalidOperationException("Brak informacji o pliku kalkulatora.");

        var values = SheetModel.Cells
            .Where(cell => cell.IsInput && !cell.IsMergedSlave)
            .OrderBy(cell => cell.Row)
            .ThenBy(cell => cell.Column)
            .Select(cell => {
                var key = (cell.Row, cell.Column);
                var value = _pendingCellValues.TryGetValue(key, out var pendingValue)
                    ? pendingValue
                    : cell.DisplayText;

                return new CalculatorVersionCellModel {
                    Row = cell.Row,
                    Column = cell.Column,
                    Value = value,
                    InputType = cell.InputType.ToString()
                };
            })
            .ToList();

        return new CalculatorVersionModel {
            FormatVersion = CalculatorVersionService.CurrentFormatVersion,
            CalculatorName = CalculatorName,
            CalculatorFileIdentity = CalculatorVersionService.CreateFileIdentity(_calculatorInfo.FilePath),
            SheetName = SheetModel.SheetName,
            CreatedAt = DateTime.Now,
            Values = values
        };
    }

    private void ValidateVersion(CalculatorVersionModel version) {
        if (_calculatorInfo == null)
            throw new InvalidOperationException("Brak informacji o pliku kalkulatora.");

        CalculatorVersionService.ValidateBelongsToCalculator(version, _calculatorInfo.FilePath);

        if (!string.IsNullOrWhiteSpace(version.CalculatorName) &&
            !string.Equals(version.CalculatorName, CalculatorName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"Plik wersji jest zapisany dla kalkulatora '{version.CalculatorName}', a aktualnie otwarty jest '{CalculatorName}'.");

        if (version.Values.Count == 0)
            throw new InvalidOperationException("Plik wersji nie zawiera żadnych zapisanych wartości.");
    }


    private List<KeyValuePair<(int Row, int Column), string>> BuildPendingValuesWithSynchronizedDropdowns(
        List<KeyValuePair<(int Row, int Column), string>> pendingValues) {
        if (SheetModel == null || pendingValues.Count == 0)
            return pendingValues;

        var result = new Dictionary<(int Row, int Column), string>();

        foreach (var pendingValue in pendingValues)
            result[pendingValue.Key] = pendingValue.Value;

        var dropdownCells = SheetModel.Cells
            .Where(cell =>
                !cell.IsMergedSlave &&
                cell.InputType == CellInputType.ComboBox &&
                cell.DropdownValues.Count > 0)
            .ToList();

        if (dropdownCells.Count < 2)
            return result
                .Select(value => new KeyValuePair<(int Row, int Column), string>(value.Key, value.Value))
                .ToList();

        var dropdownsByPosition = dropdownCells.ToDictionary(cell => (cell.Row, cell.Column), cell => cell);

        foreach (var pendingValue in pendingValues) {
            if (!dropdownsByPosition.TryGetValue(pendingValue.Key, out var changedDropdown))
                continue;

            var changedPreviousValue = GetCommittedCellValue(changedDropdown);
            var changedPreviousText = NormalizeDropdownText(changedPreviousValue);
            var selectedText = NormalizeDropdownText(pendingValue.Value);
            var changedPreviousIndex = GetCommittedDropdownSelectedIndex(changedDropdown, changedPreviousValue);
            var selectedIndex = GetDropdownSelectedIndex(changedDropdown, pendingValue.Value);

            foreach (var candidateDropdown in from candidateDropdown in dropdownCells
                     where candidateDropdown.Row != changedDropdown.Row ||
                           candidateDropdown.Column != changedDropdown.Column
                     let candidatePreviousValue = GetCommittedCellValue(candidateDropdown)
                     let candidatePreviousText = NormalizeDropdownText(candidatePreviousValue)
                     let candidatePreviousIndex =
                         GetCommittedDropdownSelectedIndex(candidateDropdown, candidatePreviousValue)
                     where ShouldSynchronizeDropdown(
                         changedDropdown,
                         candidateDropdown,
                         changedPreviousText,
                         candidatePreviousText,
                         selectedText,
                         changedPreviousIndex,
                         candidatePreviousIndex,
                         selectedIndex)
                     select candidateDropdown) {
                if (!TryCreateSynchronizedDropdownValue(
                        candidateDropdown,
                        pendingValue.Value,
                        selectedIndex,
                        out var synchronizedValue))
                    continue;

                result[(candidateDropdown.Row, candidateDropdown.Column)] = synchronizedValue;
            }
        }

        return result
            .Select(value => new KeyValuePair<(int Row, int Column), string>(value.Key, value.Value))
            .ToList();
    }

    private static bool ShouldSynchronizeDropdown(
        CellModel changedDropdown,
        CellModel candidateDropdown,
        string changedPreviousText,
        string candidatePreviousText,
        string selectedText,
        int changedPreviousIndex,
        int candidatePreviousIndex,
        int selectedIndex) {
        if (HaveSameInputTarget(changedDropdown, candidateDropdown))
            return true;

        if (HaveSameDropdownLinkedCell(changedDropdown, candidateDropdown))
            return true;

        if (HaveSameDropdownListSource(changedDropdown, candidateDropdown) &&
            CanSynchronizeBySelectedIndex(
                changedDropdown,
                candidateDropdown,
                changedPreviousIndex,
                candidatePreviousIndex,
                selectedIndex))
            return true;

        if (CanSynchronizeBySelectedIndex(
                changedDropdown,
                candidateDropdown,
                changedPreviousIndex,
                candidatePreviousIndex,
                selectedIndex))
            return true;

        if (!string.IsNullOrWhiteSpace(changedPreviousText) &&
            string.Equals(candidatePreviousText, changedPreviousText, StringComparison.OrdinalIgnoreCase))
            return true;

        if (!string.IsNullOrWhiteSpace(selectedText) &&
            string.Equals(candidatePreviousText, selectedText, StringComparison.OrdinalIgnoreCase))
            return true;

        return string.IsNullOrWhiteSpace(candidatePreviousText) &&
               string.IsNullOrWhiteSpace(changedPreviousText);
    }

    private static bool CanSynchronizeBySelectedIndex(
        CellModel changedDropdown,
        CellModel candidateDropdown,
        int changedPreviousIndex,
        int candidatePreviousIndex,
        int selectedIndex) {
        if (selectedIndex <= 0 || changedPreviousIndex <= 0 || candidatePreviousIndex <= 0)
            return false;

        if (candidatePreviousIndex != changedPreviousIndex)
            return false;

        return selectedIndex <= candidateDropdown.DropdownValues.Count;
    }

    private static bool TryCreateSynchronizedDropdownValue(
        CellModel candidateDropdown,
        string selectedValue,
        int selectedIndex,
        out string synchronizedValue) {
        if (selectedIndex > 0 && selectedIndex <= candidateDropdown.DropdownValues.Count) {
            synchronizedValue = candidateDropdown.DropdownValues[selectedIndex - 1];
            return true;
        }

        if (DropdownContainsValue(candidateDropdown, selectedValue)) {
            synchronizedValue = selectedValue;
            return true;
        }

        synchronizedValue = string.Empty;
        return false;
    }

    private static bool DropdownContainsValue(CellModel dropdown, string value) {
        var normalizedValue = NormalizeDropdownText(value);

        if (string.IsNullOrWhiteSpace(normalizedValue))
            return false;

        return dropdown.DropdownValues.Any(dropdownValue =>
            string.Equals(
                NormalizeDropdownText(dropdownValue),
                normalizedValue,
                StringComparison.OrdinalIgnoreCase));
    }

    private static bool HaveSameInputTarget(CellModel firstDropdown, CellModel secondDropdown) {
        if (firstDropdown.InputTargetRow == null ||
            firstDropdown.InputTargetColumn == null ||
            secondDropdown.InputTargetRow == null ||
            secondDropdown.InputTargetColumn == null)
            return false;

        return firstDropdown.InputTargetRow == secondDropdown.InputTargetRow &&
               firstDropdown.InputTargetColumn == secondDropdown.InputTargetColumn &&
               string.Equals(
                   firstDropdown.InputTargetSheetName ?? string.Empty,
                   secondDropdown.InputTargetSheetName ?? string.Empty,
                   StringComparison.OrdinalIgnoreCase);
    }

    private static bool HaveSameDropdownLinkedCell(CellModel firstDropdown, CellModel secondDropdown) {
        return AreDropdownReferencesEqual(
            firstDropdown.DropdownLinkedCellReference,
            secondDropdown.DropdownLinkedCellReference);
    }

    private static bool HaveSameDropdownListSource(CellModel firstDropdown, CellModel secondDropdown) {
        return AreDropdownReferencesEqual(
            firstDropdown.DropdownListSourceReference,
            secondDropdown.DropdownListSourceReference);
    }

    private static bool AreDropdownReferencesEqual(string? firstReference, string? secondReference) {
        var firstNormalizedReference = NormalizeDropdownReference(firstReference);
        var secondNormalizedReference = NormalizeDropdownReference(secondReference);

        return !string.IsNullOrWhiteSpace(firstNormalizedReference) &&
               !string.IsNullOrWhiteSpace(secondNormalizedReference) &&
               string.Equals(firstNormalizedReference, secondNormalizedReference, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeDropdownReference(string? reference) {
        if (string.IsNullOrWhiteSpace(reference))
            return string.Empty;

        var normalizedReference = reference.Trim();

        if (normalizedReference.StartsWith('='))
            normalizedReference = normalizedReference[1..].Trim();

        return normalizedReference.Replace("$", string.Empty).Replace("''", "'").Trim();
    }

    private void WritePendingValues(
        IEnumerable<KeyValuePair<(int Row, int Column), string>> pendingValues) {
        var pendingValuesList = pendingValues.ToList();
        var cellModels = SheetModel?.Cells
                             .Where(cell => !cell.IsMergedSlave)
                             .ToDictionary(cell => (cell.Row, cell.Column), cell => cell)
                         ?? [];
        var writtenDropdowns = new List<PendingDropdownWrite>();

        foreach (var pendingValue in pendingValuesList) {
            cellModels.TryGetValue(pendingValue.Key, out var cellModel);

            var previousValue = cellModel?.InputType == CellInputType.ComboBox
                ? GetCommittedCellValue(cellModel)
                : string.Empty;

            var previousSelectedIndex = cellModel?.InputType == CellInputType.ComboBox
                ? GetCommittedDropdownSelectedIndex(cellModel, previousValue)
                : 0;

            var selectedIndex = cellModel?.InputType == CellInputType.ComboBox
                ? GetDropdownSelectedIndex(cellModel, pendingValue.Value)
                : 0;

            var dropdownControlWasWritten = cellModel?.InputType == CellInputType.ComboBox &&
                                            _excelSession.TryWriteDropdownControlValue(
                                                cellModel,
                                                pendingValue.Value,
                                                selectedIndex,
                                                false);

            if (dropdownControlWasWritten && cellModel?.InputType == CellInputType.ComboBox) {
                _excelSession.SynchronizeDropdownControlsByPreviousValue(
                    cellModel,
                    previousValue,
                    previousSelectedIndex,
                    pendingValue.Value,
                    selectedIndex);

                writtenDropdowns.Add(new PendingDropdownWrite(
                    cellModel,
                    previousValue,
                    previousSelectedIndex,
                    pendingValue.Value,
                    selectedIndex));
            }

            var hasExplicitInputTarget = cellModel is { InputTargetRow: not null, InputTargetColumn: not null };

            switch (dropdownControlWasWritten) {
                case true when cellModel?.DropdownWritesSelectedIndex == true:
                case true when !hasExplicitInputTarget:
                    continue;
            }

            if (cellModel?.InputType == CellInputType.ComboBox &&
                cellModel.DropdownWritesSelectedIndex &&
                selectedIndex <= 0)
                continue;

            var targetRow = cellModel?.InputTargetRow ?? pendingValue.Key.Row;
            var targetColumn = cellModel?.InputTargetColumn ?? pendingValue.Key.Column;
            var targetSheetName = cellModel?.InputTargetSheetName;
            var value = CreateExcelInputValue(cellModel, pendingValue.Value, selectedIndex);

            _excelSession.WriteCellValue(
                targetRow,
                targetColumn,
                value,
                targetSheetName);
        }

        RunPendingDropdownMacros(writtenDropdowns);
    }

    private void RunPendingDropdownMacros(List<PendingDropdownWrite> writtenDropdowns) {
        if (writtenDropdowns.Count == 0)
            return;

        var executedDropdownKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var writtenDropdown in from writtenDropdown in writtenDropdowns
                 let dropdownKey = CreateDropdownMacroExecutionKey(writtenDropdown.CellModel)
                 where executedDropdownKeys.Add(dropdownKey)
                 where _excelSession.TryRunDropdownControlAssignedMacro(
                     writtenDropdown.CellModel,
                     writtenDropdown.SelectedValue,
                     writtenDropdown.SelectedIndex)
                 select writtenDropdown) {
            _excelSession.TryWriteDropdownControlValue(
                writtenDropdown.CellModel,
                writtenDropdown.SelectedValue,
                writtenDropdown.SelectedIndex,
                false);

            _excelSession.SynchronizeDropdownControlsByPreviousValue(
                writtenDropdown.CellModel,
                writtenDropdown.PreviousValue,
                writtenDropdown.PreviousSelectedIndex,
                writtenDropdown.SelectedValue,
                writtenDropdown.SelectedIndex);
        }
    }

    private static string CreateDropdownMacroExecutionKey(CellModel cellModel) {
        var linkedCellReference = NormalizeDropdownReference(cellModel.DropdownLinkedCellReference);

        if (!string.IsNullOrWhiteSpace(linkedCellReference))
            return $"LinkedCell:{linkedCellReference}";

        if (cellModel is { InputTargetRow: not null, InputTargetColumn: not null })
            return
                $"InputTarget:{cellModel.InputTargetSheetName ?? string.Empty}:{cellModel.InputTargetRow}:{cellModel.InputTargetColumn}";

        return !string.IsNullOrWhiteSpace(cellModel.DropdownControlName)
            ? $"Control:{cellModel.DropdownControlName}"
            : $"Cell:{cellModel.Row}:{cellModel.Column}";
    }

    private static object? CreateExcelInputValue(CellModel? cellModel, string text, int selectedIndex = 0) {
        if (cellModel?.InputType == CellInputType.ComboBox && cellModel.DropdownWritesSelectedIndex) {
            var resolvedSelectedIndex = selectedIndex > 0
                ? selectedIndex
                : GetDropdownSelectedIndex(cellModel, text);

            if (resolvedSelectedIndex > 0)
                return resolvedSelectedIndex;
        }

        return TryParseNumeric(text, out var number)
            ? number
            : string.IsNullOrWhiteSpace(text)
                ? null
                : text;
    }

    private static int GetCommittedDropdownSelectedIndex(CellModel cellModel, string committedValue) {
        if (cellModel.DropdownSelectedIndex is > 0)
            return cellModel.DropdownSelectedIndex.Value;

        return GetDropdownSelectedIndex(cellModel, committedValue);
    }

    private static int GetDropdownSelectedIndex(CellModel cellModel, string selectedValue) {
        if (cellModel.DropdownValues.Count == 0 || string.IsNullOrWhiteSpace(selectedValue))
            return 0;

        if (cellModel.DropdownWritesSelectedIndex &&
            int.TryParse(selectedValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture,
                out var selectedNumericIndex) &&
            selectedNumericIndex > 0 &&
            selectedNumericIndex <= cellModel.DropdownValues.Count)
            return selectedNumericIndex;

        var normalizedSelectedValue = NormalizeDropdownText(selectedValue);

        for (var index = 0; index < cellModel.DropdownValues.Count; index++)
            if (AreDropdownValuesEqual(cellModel.DropdownValues[index], selectedValue, normalizedSelectedValue))
                return index + 1;

        return 0;
    }

    private static bool AreDropdownValuesEqual(
        string? dropdownValue,
        string selectedValue,
        string normalizedSelectedValue) {
        var normalizedDropdownValue = NormalizeDropdownText(dropdownValue);

        if (string.Equals(normalizedDropdownValue, normalizedSelectedValue, StringComparison.OrdinalIgnoreCase))
            return true;

        if (string.Equals(
                normalizedDropdownValue.Replace(" ", string.Empty),
                normalizedSelectedValue.Replace(" ", string.Empty),
                StringComparison.OrdinalIgnoreCase))
            return true;

        return !string.IsNullOrWhiteSpace(dropdownValue) &&
               TryParseNumeric(dropdownValue, out var dropdownNumber) &&
               TryParseNumeric(selectedValue, out var selectedNumber) &&
               Math.Abs(dropdownNumber - selectedNumber) < 0.0000001;
    }

    private static string NormalizeDropdownText(string? value) {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var text = value.Replace('\u00A0', ' ');
        var parts = text.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        return string.Join(" ", parts).Trim();
    }


    private string GetCommittedCellValue(CellModel cellModel) {
        return _committedCellValues.TryGetValue((cellModel.Row, cellModel.Column), out var committedValue)
            ? committedValue
            : cellModel.DisplayText;
    }

    private void UpdateCommittedCellValues(SheetModel model) {
        _committedCellValues.Clear();

        foreach (var cell in model.Cells.Where(cell => !cell.IsMergedSlave))
            _committedCellValues[(cell.Row, cell.Column)] = cell.DisplayText;
    }

    private void ClearPerformanceMessage() {
        _operationPerformanceMessage = string.Empty;
        PerformanceMessage = string.Empty;
    }

    private void SetOperationPerformanceMessage(string message) {
        _operationPerformanceMessage = message;
        PerformanceMessage = message;
    }

    private static string FormatDuration(TimeSpan duration) {
        return duration.TotalSeconds >= 1.0
            ? $"{duration.TotalSeconds:N2} s"
            : $"{duration.TotalMilliseconds:N0} ms";
    }

    private void ClearPendingChanges() {
        _pendingCellValues.Clear();
        HasPendingChanges = false;
    }

    private static bool TryParseNumeric(string text, out double result) {
        if (string.IsNullOrWhiteSpace(text)) {
            result = 0;
            return false;
        }

        var normalizedValue = text.Trim().Replace(',', '.');

        return double.TryParse(
            normalizedValue,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out result);
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

    private void OnPropertyChanged([CallerMemberName] string? name = null) {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    private sealed record PendingDropdownWrite(
        CellModel CellModel,
        string PreviousValue,
        int PreviousSelectedIndex,
        string SelectedValue,
        int SelectedIndex);
}