using System.Collections.ObjectModel;
using System.ComponentModel;
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

        var operationName = "tworzenia kopii roboczej";

        try {
            if (!File.Exists(calculatorInfo.FilePath))
                throw new FileNotFoundException("Nie znaleziono wskazanego pliku kalkulatora.",
                    calculatorInfo.FilePath);

            var workingPath = _workingCopy.CreateWorkingCopy(calculatorInfo.FilePath);

            operationName = "uruchamiania programu Excel i otwierania skoroszytu";
            StatusMessage = "Uruchamianie Excela…";
            await _worker.InvokeAsync(() => _excelSession.OpenSession(workingPath));

            operationName = "odczytu pierwszego arkusza";
            SheetModel model;

            if (_sheetLayoutCache.TryLoad(calculatorInfo, out var cachedModel) && cachedModel != null) {
                StatusMessage = "Odczyt wartości z arkusza…";
                model = await _worker.InvokeAsync(() => _sheetReader.RefreshCellValues(_excelSession, cachedModel));
            }
            else {
                StatusMessage = "Odczyt układu arkusza…";
                model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));
                _sheetLayoutCache.TrySave(calculatorInfo, model);
            }

            operationName = "wczytywania konfiguracji makr";
            var buttons = MacroConfigService.LoadForCalculator(calculatorInfo.FilePath);

            if (_disposed) return;

            ClearPendingChanges();
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

    public async Task SaveVersionAsync(string versionFilePath) {
        if (_disposed || IsLoading) return;

        try {
            ErrorMessage = string.Empty;

            if (SheetModel == null)
                throw new InvalidOperationException("Nie ma wczytanego arkusza, więc nie można zapisać wersji.");

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

        try {
            if (SheetModel == null)
                throw new InvalidOperationException("Nie ma wczytanego arkusza, więc nie można wczytać wersji.");

            var version = CalculatorVersionService.Load(versionFilePath);
            ValidateVersion(version);

            var values = version.Values
                .Where(value => value.Row > 0 && value.Column > 0)
                .Select(value => new KeyValuePair<(int Row, int Column), string>(
                    (value.Row, value.Column),
                    value.Value))
                .ToList();

            await _worker.InvokeAsync(() => {
                WriteCellValues(values);
                _excelSession.Recalculate();
            });

            var model = await RefreshCurrentSheetModelAsync();

            if (_disposed) return;

            ClearPendingChanges();
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

    private async Task CalculateAsync() {
        if (_disposed || IsLoading || HasError) return;

        IsLoading = true;
        ErrorMessage = string.Empty;
        StatusMessage = "Przeliczanie…";

        var pendingValues = _pendingCellValues.ToList();

        try {
            await _worker.InvokeAsync(() => {
                WriteCellValues(pendingValues);
                _excelSession.Recalculate();
            });

            var model = await RefreshCurrentSheetModelAsync();

            if (_disposed) return;

            ClearPendingChanges();
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

        var pendingValues = _pendingCellValues.ToList();

        try {
            await _worker.InvokeAsync(() => {
                WriteCellValues(pendingValues);
                _excelSession.RunMacro(config.MacroName);
            });

            var model = config.RefreshLayoutAfterRun
                ? await ReadFullSheetModelAndUpdateCacheAsync()
                : await RefreshCurrentSheetModelAsync();

            if (_disposed) return;

            ClearPendingChanges();
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

    private async Task<SheetModel> RefreshCurrentSheetModelAsync() {
        if (SheetModel == null)
            throw new InvalidOperationException("Nie ma wczytanego arkusza.");

        return await _worker.InvokeAsync(() => _sheetReader.RefreshCellValues(_excelSession, SheetModel));
    }

    private async Task<SheetModel> ReadFullSheetModelAndUpdateCacheAsync() {
        var model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));

        if (_calculatorInfo != null)
            _sheetLayoutCache.TrySave(_calculatorInfo, model);

        return model;
    }

    private CalculatorVersionModel CreateVersionModel() {
        if (SheetModel == null)
            throw new InvalidOperationException("Nie ma wczytanego arkusza.");

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
            FormatVersion = 1,
            CalculatorName = CalculatorName,
            SheetName = SheetModel.SheetName,
            CreatedAt = DateTime.Now,
            Values = values
        };
    }

    private void ValidateVersion(CalculatorVersionModel version) {
        if (!string.IsNullOrWhiteSpace(version.CalculatorName) &&
            !string.Equals(version.CalculatorName, CalculatorName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"Plik wersji jest zapisany dla kalkulatora '{version.CalculatorName}', a aktualnie otwarty jest '{CalculatorName}'.");

        if (version.Values.Count == 0)
            throw new InvalidOperationException("Plik wersji nie zawiera żadnych zapisanych wartości.");
    }

    private void WriteCellValues(
        IEnumerable<KeyValuePair<(int Row, int Column), string>> values) {
        foreach (var valuePair in values) {
            var value = TryParseNumeric(valuePair.Value, out var number)
                ? number
                : (object?)valuePair.Value;

            _excelSession.WriteCellValue(
                valuePair.Key.Row,
                valuePair.Key.Column,
                value);
        }
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