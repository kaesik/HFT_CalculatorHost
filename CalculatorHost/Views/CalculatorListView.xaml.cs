using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CalculatorHost.Models;
using CalculatorHost.ViewModels;

namespace CalculatorHost.Views;

public partial class CalculatorListView {
    private readonly MainViewModel _viewModel;
    private bool _isSubscribed;

    public CalculatorListView(MainViewModel viewModel) {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;

        CalculatorList.ItemsSource = viewModel.Calculators;
        FolderPathText.Text = viewModel.FolderPath;
        MissingFolderPath.Text = viewModel.FolderPath;
        UpdateCount();

        Loaded += CalculatorListView_Loaded;
        Unloaded += CalculatorListView_Unloaded;
    }

    private void CalculatorListView_Loaded(object sender, RoutedEventArgs e) {
        if (_isSubscribed) return;

        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
        _isSubscribed = true;

        CalculatorList.ItemsSource = _viewModel.Calculators;
        FolderPathText.Text = _viewModel.FolderPath;
        MissingFolderPath.Text = _viewModel.FolderPath;
        UpdateCount();
    }

    private void CalculatorListView_Unloaded(object sender, RoutedEventArgs e) {
        if (!_isSubscribed) return;

        _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        _isSubscribed = false;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e) {
        switch (e.PropertyName) {
            case nameof(MainViewModel.Calculators):
                CalculatorList.ItemsSource = _viewModel.Calculators;
                UpdateCount();
                break;
            case nameof(MainViewModel.FolderPath):
                FolderPathText.Text = _viewModel.FolderPath;
                MissingFolderPath.Text = _viewModel.FolderPath;
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
}