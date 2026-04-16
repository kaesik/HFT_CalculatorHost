using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using CalculatorHost.Models;
using CalculatorHost.Services;

namespace CalculatorHost.ViewModels;

public class MainViewModel : INotifyPropertyChanged {
    private readonly CalculatorScannerService _scanner;

    private ObservableCollection<CalculatorInfo> _calculators = [];
    private string _errorMessage = string.Empty;
    private readonly string _folderPath = string.Empty;
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
            _selectedCalculator = value;
            OnPropertyChanged();
        }
    }

    public string FolderPath {
        get => _folderPath;
        private init {
            _folderPath = value;
            OnPropertyChanged();
        }
    }

    private bool IsFolderMissing {
        get => _isFolderMissing;
        set {
            _isFolderMissing = value;
            OnPropertyChanged();
        }
    }

    public string ErrorMessage {
        get => _errorMessage;
        set {
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
            IsFolderMissing = !_scanner.FolderExists();
            ErrorMessage = string.Empty;

            var found = _scanner.ScanFolder();
            Calculators = new ObservableCollection<CalculatorInfo>(found);

            if (!IsFolderMissing && found.Count == 0)
                ErrorMessage = $"Folder '{FolderPath}' jest pusty — brak plików .xlsm.";
        }
        catch (Exception ex) {
            ErrorMessage = $"Błąd podczas skanowania folderu: {ex.Message}";
        }
    }

    private async Task OpenSelectedCalculatorAsync() {
        if (SelectedCalculator == null) return;
        CalculatorOpenRequested?.Invoke(SelectedCalculator);
        await Task.CompletedTask;
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null) {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}