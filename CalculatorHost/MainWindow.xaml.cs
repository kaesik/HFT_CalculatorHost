using System.Windows;
using System.Windows.Input;
using CalculatorHost.Models;
using CalculatorHost.Services;
using CalculatorHost.ViewModels;
using CalculatorHost.Views;

namespace CalculatorHost;

public partial class MainWindow {
    private readonly MainViewModel _mainViewModel;
    private CalculatorViewModel? _activeCalculatorViewModel;

    public MainWindow() {
        InitializeComponent();

        var scanner = new CalculatorScannerService();
        var workingCopyForCleanup = new WorkingCopyService();
        workingCopyForCleanup.CleanAllOrphanedSessions();

        _mainViewModel = new MainViewModel(scanner);
        _mainViewModel.CalculatorOpenRequested += OnCalculatorOpenRequested;

        ShowListView();
    }

    private void ShowListView() {
        TitleCalculatorName.Text = string.Empty;
        DisposeActiveCalculator();
        MainContent.Content = new CalculatorListView(_mainViewModel);
    }

    private async void OnCalculatorOpenRequested(CalculatorInfo calculatorInfo) {
        try {
            DisposeActiveCalculator();

            var session = new ExcelSessionService();
            var reader = new SheetReaderService();
            var macroConfig = new MacroConfigService();
            var workingCopy = new WorkingCopyService();

            _activeCalculatorViewModel = new CalculatorViewModel(
                session,
                reader,
                macroConfig,
                workingCopy,
                App.ExcelWorker);

            _activeCalculatorViewModel.CloseRequested += OnCalculatorCloseRequested;

            var calculatorView = new CalculatorView(_activeCalculatorViewModel);
            MainContent.Content = calculatorView;
            TitleCalculatorName.Text = $"— {calculatorInfo.DisplayName}";

            await _activeCalculatorViewModel.LoadCalculatorAsync(calculatorInfo);
        }
        catch (Exception exception) {
            if (_activeCalculatorViewModel != null) {
                _activeCalculatorViewModel.ShowExternalError(
                    "Nie udało się uruchomić widoku kalkulatora",
                    exception);

                return;
            }

            MessageBox.Show(
                $"Nie udało się uruchomić kalkulatora: {exception.Message}",
                "Błąd uruchamiania kalkulatora",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            ShowListView();
        }
    }

    private void OnCalculatorCloseRequested() {
        ShowListView();
    }

    private void DisposeActiveCalculator() {
        if (_activeCalculatorViewModel == null) return;

        _activeCalculatorViewModel.CloseRequested -= OnCalculatorCloseRequested;
        _activeCalculatorViewModel.Dispose();
        _activeCalculatorViewModel = null;
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
        if (e.ClickCount == 2)
            WindowState = WindowState == WindowState.Maximized
                ? WindowState.Normal
                : WindowState.Maximized;
        else
            try {
                DragMove();
            }
            catch {
            }
    }

    private void ThemeToggle_Click(object sender, RoutedEventArgs e) {
        ThemeService.Toggle();
    }

    private void Minimize_Click(object sender, RoutedEventArgs e) {
        WindowState = WindowState.Minimized;
    }

    private void Close_Click(object sender, RoutedEventArgs e) {
        Close();
    }

    protected override void OnClosed(EventArgs e) {
        _mainViewModel.CalculatorOpenRequested -= OnCalculatorOpenRequested;
        DisposeActiveCalculator();
        base.OnClosed(e);
    }
}