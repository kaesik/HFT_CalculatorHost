using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using CalculatorHost.Models;
using CalculatorHost.Services;

namespace CalculatorHost.ViewModels;

public class CalculatorViewModel : INotifyPropertyChanged, IDisposable {
    private readonly ExcelSessionService _excelSession;
    private readonly MacroConfigService _macroConfig;
    private readonly SheetReaderService _sheetReader;
    private readonly ExcelWorker _worker;
    private readonly WorkingCopyService _workingCopy;
    private string _calculatorName = string.Empty;
    private bool _disposed;
    private string _errorMessage = string.Empty;
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
        _macroConfig = macroConfig;
        _workingCopy = workingCopy;
        _worker = worker;

        RunMacroCommand = new AsyncRelayCommand<MacroButtonConfig>(RunMacroAsync, _ => !IsLoading);
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

    private bool IsLoading {
        get => _isLoading;
        set {
            _isLoading = value;
            OnPropertyChanged();
        }
    }

    private string ErrorMessage {
        get => _errorMessage;
        set {
            _errorMessage = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasError));
        }
    }

    public bool HasError => !string.IsNullOrEmpty(ErrorMessage);

    public string CalculatorName {
        get => _calculatorName;
        private set {
            _calculatorName = value;
            OnPropertyChanged();
        }
    }

    public string StatusMessage {
        get => _statusMessage;
        set {
            _statusMessage = value;
            OnPropertyChanged();
        }
    }

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
            /* Best-effort cleanup */
        }

        _workingCopy.CleanCurrentSession();
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public event Action? CloseRequested;

    public async Task LoadCalculatorAsync(CalculatorInfo calculatorInfo) {
        CalculatorName = calculatorInfo.DisplayName;
        IsLoading = true;
        ErrorMessage = string.Empty;
        StatusMessage = "Tworzenie kopii roboczej…";

        try {
            var workingPath = _workingCopy.CreateWorkingCopy(calculatorInfo.FilePath);
            StatusMessage = "Uruchamianie Excela…";

            await _worker.InvokeAsync(() => _excelSession.OpenSession(workingPath));

            StatusMessage = "Odczyt arkusza…";
            var model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));
            var buttons = MacroConfigService.LoadForCalculator(calculatorInfo.FilePath);

            Application.Current.Dispatcher.Invoke(() => {
                SheetModel = model;
                MacroButtons = new ObservableCollection<MacroButtonConfig>(buttons);
                StatusMessage = string.Empty;
            });
        }
        catch (Exception ex) {
            Application.Current.Dispatcher.Invoke(() => {
                ErrorMessage = FormatExceptionMessage("Nie udało się otworzyć kalkulatora", ex);
                StatusMessage = string.Empty;
            });
        }
        finally {
            Application.Current.Dispatcher.Invoke(() => IsLoading = false);
        }
    }

    public async Task UpdateCellValueAsync(int row, int column, string value) {
        if (IsLoading) return;
        IsLoading = true;
        StatusMessage = "Przeliczanie…";
        ErrorMessage = string.Empty;

        try {
            var parsedValue = TryParseNumeric(value, out var number) ? number : (object?)value;
            await _worker.InvokeAsync(() => _excelSession.WriteCellValue(row, column, parsedValue));

            var model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));

            Application.Current.Dispatcher.Invoke(() => {
                SheetModel = model;
                StatusMessage = string.Empty;
            });
        }
        catch (Exception ex) {
            Application.Current.Dispatcher.Invoke(() => {
                ErrorMessage = FormatExceptionMessage("Błąd zapisu wartości", ex);
                StatusMessage = string.Empty;
            });
        }
        finally {
            Application.Current.Dispatcher.Invoke(() => IsLoading = false);
        }
    }

    private async Task RunMacroAsync(MacroButtonConfig? config) {
        if (config == null || IsLoading) return;
        IsLoading = true;
        ErrorMessage = string.Empty;
        StatusMessage = $"Wykonywanie: {config.Label}…";

        try {
            await _worker.InvokeAsync(() => _excelSession.RunMacro(config.MacroName));

            var model = await _worker.InvokeAsync(() => _sheetReader.ReadFirstSheet(_excelSession));

            Application.Current.Dispatcher.Invoke(() => {
                SheetModel = model;
                StatusMessage = string.Empty;
            });
        }
        catch (Exception ex) {
            Application.Current.Dispatcher.Invoke(() => {
                ErrorMessage = FormatExceptionMessage($"Makro '{config.Label}' nie powiodło się", ex);
                StatusMessage = string.Empty;
            });
        }
        finally {
            Application.Current.Dispatcher.Invoke(() => IsLoading = false);
        }
    }

    public void RequestClose() {
        CloseRequested?.Invoke();
    }

    private static bool TryParseNumeric(string text, out double result) {
        if (string.IsNullOrWhiteSpace(text)) {
            result = 0;
            return false;
        }

        if (double.TryParse(text, NumberStyles.Any,
                CultureInfo.InvariantCulture, out result)) return true;

        return double.TryParse(text.Replace(',', '.'), NumberStyles.Any,
            CultureInfo.InvariantCulture, out result);
    }

    private static string FormatExceptionMessage(string prefix, Exception ex) {
        var inner = ex.InnerException?.Message;
        return inner != null
            ? $"{prefix}: {ex.Message}\n({inner})"
            : $"{prefix}: {ex.Message}";
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null) {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}