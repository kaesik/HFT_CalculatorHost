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
    private readonly SheetReaderService _sheetReader;
    private readonly ExcelWorker _worker;
    private readonly WorkingCopyService _workingCopy;
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
        ExcelWorker worker) {
        _excelSession = excelSession;
        _sheetReader = sheetReader;
        _workingCopy = workingCopy;
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
            StatusMessage = "Odczyt arkusza…";
            var model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));

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
                WritePendingValues(pendingValues);
                _excelSession.Recalculate();
            });

            var model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));

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
                WritePendingValues(pendingValues);
                _excelSession.RunMacro(config.MacroName);
            });

            var model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));

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