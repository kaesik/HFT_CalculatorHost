using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using CalculatorHost.Models;
using CalculatorHost.Services;

namespace CalculatorHost.ViewModels;

public class MainViewModel : INotifyPropertyChanged {
    private readonly CalculatorScannerService _scanner;
    private ObservableCollection<CalculatorInfo> _calculators = [];
    private string _errorMessage = string.Empty;
    private string _folderPath = string.Empty;
    private bool _isFolderMissing;
    private CalculatorInfo? _selectedCalculator;

    public MainViewModel(CalculatorScannerService scanner) {
        _scanner = scanner;
        FolderPath = _scanner.CalculatorsFolder;
        RefreshCommand = new RelayCommand(LoadCalculators);
        OpenCalculatorCommand = new AsyncRelayCommand(OpenSelectedCalculatorAsync, () => SelectedCalculator != null);
        LoadCalculators();
    }

    public ObservableCollection<CalculatorInfo> Calculators {
        get => _calculators;
        private set {
            _calculators = value;
            OnPropertyChanged();
        }
    }

    public CalculatorInfo? SelectedCalculator {
        get => _selectedCalculator;
        set {
            if (Equals(_selectedCalculator, value)) return;

            _selectedCalculator = value;
            OnPropertyChanged();
            CommandManager.InvalidateRequerySuggested();
        }
    }

    public string FolderPath {
        get => _folderPath;
        private set {
            if (_folderPath == value) return;

            _folderPath = value;
            OnPropertyChanged();
        }
    }

    public bool IsFolderMissing {
        get => _isFolderMissing;
        private set {
            if (_isFolderMissing == value) return;

            _isFolderMissing = value;
            OnPropertyChanged();
        }
    }

    public string ErrorMessage {
        get => _errorMessage;
        private set {
            if (_errorMessage == value) return;

            _errorMessage = value;
            OnPropertyChanged();
        }
    }

    public RelayCommand RefreshCommand { get; }
    public AsyncRelayCommand OpenCalculatorCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action<CalculatorInfo>? CalculatorOpenRequested;

    private void LoadCalculators() {
        try {
            ErrorMessage = string.Empty;
            IsFolderMissing = !_scanner.FolderExists();
            SelectedCalculator = null;

            if (IsFolderMissing) {
                Calculators = [];
                return;
            }

            var found = _scanner.ScanFolder();
            Calculators = new ObservableCollection<CalculatorInfo>(found);

            if (found.Count == 0)
                ErrorMessage = $"Folder '{FolderPath}' jest pusty - brak plików .xlsm.";
        }
        catch (Exception exception) {
            IsFolderMissing = false;
            SelectedCalculator = null;
            Calculators = [];
            ErrorMessage = $"Błąd podczas skanowania folderu: {exception.Message}";
        }
    }

    private Task OpenSelectedCalculatorAsync() {
        if (SelectedCalculator != null)
            CalculatorOpenRequested?.Invoke(SelectedCalculator);

        return Task.CompletedTask;
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null) {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}