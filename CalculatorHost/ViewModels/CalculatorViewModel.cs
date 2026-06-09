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

public class CalculatorViewModel : INotifyPropertyChanged, IDisposable {
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
        MacroConfigService macroConfig,
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

        var originalValue = SheetModel.Cells
            .FirstOrDefault(cell => cell.Row == row && cell.Column == column)
            ?.DisplayText ?? string.Empty;

        if (string.Equals(originalValue, value, StringComparison.Ordinal))
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
            if (SheetModel == null)
                throw new InvalidOperationException("Nie ma wczytanego arkusza, więc nie można wczytać wersji.");

            if (_calculatorInfo == null)
                throw new InvalidOperationException("Brak informacji o pliku kalkulatora.");

            var version = CalculatorVersionService.Load(versionFilePath);
            ValidateVersion(version);

            var values = version.Values
                .Where(value => value.Row > 0 && value.Column > 0)
                .Select(value => new KeyValuePair<(int Row, int Column), string>(
                    (value.Row, value.Column),
                    value.Value))
                .ToList();

            var applyingStopwatch = Stopwatch.StartNew();
            await _worker.InvokeAsync(() => {
                WritePendingValues(values);
                _excelSession.Recalculate();
            });
            applyingStopwatch.Stop();

            var currentModel = SheetModel;

            if (currentModel == null)
                return;

            var refreshStopwatch = Stopwatch.StartNew();
            var model = await _worker.InvokeAsync(() => _sheetReader.RefreshCellValues(_excelSession, currentModel));
            refreshStopwatch.Stop();

            if (_disposed) return;

            ClearPendingChanges();
            SetOperationPerformanceMessage(
                $"Wczytanie wersji: {FormatDuration(applyingStopwatch.Elapsed)} · " +
                $"Odświeżenie wartości: {FormatDuration(refreshStopwatch.Elapsed)}");
            SheetModel = model;
            StatusMessage = $"Wczytano wersję: {Path.GetFileName(versionFilePath)}";
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

        var pendingValues = _pendingCellValues.ToList();

        try {
            var recalculationStopwatch = Stopwatch.StartNew();
            await _worker.InvokeAsync(() => {
                WritePendingValues(pendingValues);
                _excelSession.Recalculate();
            });
            recalculationStopwatch.Stop();

            var currentModel = SheetModel;

            if (currentModel == null)
                return;

            var refreshStopwatch = Stopwatch.StartNew();
            var model = await _worker.InvokeAsync(() => _sheetReader.RefreshCellValues(_excelSession, currentModel));
            refreshStopwatch.Stop();

            if (_disposed) return;

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

        var pendingValues = _pendingCellValues.ToList();

        try {
            var macroStopwatch = Stopwatch.StartNew();
            await _worker.InvokeAsync(() => {
                WritePendingValues(pendingValues);
                _excelSession.RunMacroButton(config);
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

    private void WritePendingValues(
        IEnumerable<KeyValuePair<(int Row, int Column), string>> pendingValues) {
        foreach (var pendingValue in pendingValues) {
            var value = TryParseNumeric(pendingValue.Value, out var number)
                ? number
                : (object?)pendingValue.Value;

            _excelSession.WriteCellValue(
                pendingValue.Key.Row,
                pendingValue.Key.Column,
                value);
        }
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
}