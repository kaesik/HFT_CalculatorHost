using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CalculatorHost.Models;
using CalculatorHost.ViewModels;

namespace CalculatorHost.Views;

public partial class CalculatorListView {
    private readonly MainViewModel _viewModel;

    public CalculatorListView(MainViewModel viewModel) {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;

        CalculatorList.ItemsSource = viewModel.Calculators;
        FolderPathText.Text = viewModel.FolderPath;
        MissingFolderPath.Text = viewModel.FolderPath;

        UpdateCount();

        viewModel.PropertyChanged += (_, e) => {
            switch (e.PropertyName) {
                case nameof(MainViewModel.Calculators):
                    CalculatorList.ItemsSource = viewModel.Calculators;
                    UpdateCount();
                    break;
                case nameof(MainViewModel.FolderPath):
                    FolderPathText.Text = viewModel.FolderPath;
                    MissingFolderPath.Text = viewModel.FolderPath;
                    break;
            }
        };
    }

    private void UpdateCount() {
        var count = _viewModel.Calculators.Count;
        CountText.Text = count switch {
            0 => "Brak kalkulatorów",
            1 => "1 kalkulator",
            _ => $"{count} kalkulatorów"
        };
    }

    private void CalculatorList_SelectionChanged(object sender, SelectionChangedEventArgs e) {
        _viewModel.SelectedCalculator = CalculatorList.SelectedItem as CalculatorInfo;
    }

    private void CalculatorList_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
        if (_viewModel.SelectedCalculator == null) return;
        if (_viewModel.OpenCalculatorCommand.CanExecute(null))
            _viewModel.OpenCalculatorCommand.Execute(null);
    }

    private void CalculatorList_KeyDown(object sender, KeyEventArgs e) {
        if (e.Key != Key.Return && e.Key != Key.Enter) return;
        if (_viewModel.SelectedCalculator == null) return;
        if (_viewModel.OpenCalculatorCommand.CanExecute(null))
            _viewModel.OpenCalculatorCommand.Execute(null);
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e) {
        CalculatorList.ItemsSource = null;
        _viewModel.RefreshCommand.Execute(null);
        CalculatorList.ItemsSource = _viewModel.Calculators;
        UpdateCount();
    }
}