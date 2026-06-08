using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using CalculatorHost.Models;
using CalculatorHost.Rendering;
using CalculatorHost.ViewModels;
using Microsoft.Win32;

namespace CalculatorHost.Views;

public partial class CalculatorView {
    private readonly CalculatorViewModel _viewModel;
    private RenderedSheet? _renderedSheet;

    public CalculatorView(CalculatorViewModel viewModel) {
        InitializeComponent();

        _viewModel = viewModel;
        DataContext = viewModel;

        viewModel.PropertyChanged += OnViewModelPropertyChanged;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs eventArguments) {
        switch (eventArguments.PropertyName) {
            case nameof(CalculatorViewModel.SheetModel):
                RenderSheet();
                break;
            case nameof(CalculatorViewModel.MacroButtons):
                RebuildMacroButtons();
                break;
        }
    }

    private void RenderSheet() {
        if (_viewModel.SheetModel == null) {
            _renderedSheet = null;
            SheetScrollViewer.Content = null;
            return;
        }

        var renderedSheet = SheetRenderer.RenderSheet(
            _viewModel.SheetModel,
            OnInputChanged,
            OnMacroButtonClicked);

        _renderedSheet = renderedSheet;
        SheetScrollViewer.Content = renderedSheet.Canvas;

        var inputCount = _viewModel.SheetModel.Cells.Count(cell => cell.IsInput);
        var totalCells = _viewModel.SheetModel.Cells.Count(cell => !cell.IsMergedSlave);

        SheetInfoText.Text =
            $"Arkusz: {_viewModel.SheetModel.SheetName}  ·  " +
            $"Wiersze: {_viewModel.SheetModel.MaxRow}  ·  " +
            $"Kolumny: {_viewModel.SheetModel.MaxColumn}  ·  " +
            $"Komórek: {totalCells}  ·  " +
            $"Pól edytowalnych: {inputCount}";
    }

    private void RebuildMacroButtons() {
        MacroButtonsPanel.Children.Clear();

        if (_viewModel.MacroButtons.Count == 0) {
            MacroButtonsBorder.Visibility = Visibility.Collapsed;
            return;
        }

        MacroButtonsBorder.Visibility = Visibility.Visible;

        foreach (var macroConfig in _viewModel.MacroButtons) {
            var tooltip = string.IsNullOrWhiteSpace(macroConfig.Tooltip)
                ? $"Uruchamia makro: {macroConfig.MacroName}"
                : macroConfig.Tooltip;

            var button = new Button {
                Style = Application.Current.FindResource("MacroButton") as Style,
                Content = macroConfig.Label,
                ToolTip = tooltip,
                Command = _viewModel.RunMacroCommand,
                CommandParameter = macroConfig
            };

            MacroButtonsPanel.Children.Add(button);
        }
    }

    private void OnInputChanged(int row, int column, string value) {
        _viewModel.SetPendingCellValue(row, column, value);
    }

    private void OnMacroButtonClicked(MacroButtonConfig macroConfig) {
        if (_viewModel.RunMacroCommand.CanExecute(macroConfig))
            _viewModel.RunMacroCommand.Execute(macroConfig);
    }

    private async void SaveVersionButton_Click(object sender, RoutedEventArgs eventArguments) {
        var dialog = new SaveFileDialog {
            Title = "Zapisz wersję kalkulatora",
            Filter = "Wersja kalkulatora (*.json)|*.json|Wszystkie pliki (*.*)|*.*",
            DefaultExt = ".json",
            AddExtension = true,
            OverwritePrompt = true,
            FileName = CreateDefaultVersionFileName()
        };

        if (dialog.ShowDialog() != true) return;

        await _viewModel.SaveVersionAsync(dialog.FileName);
    }

    private async void LoadVersionButton_Click(object sender, RoutedEventArgs eventArguments) {
        var dialog = new OpenFileDialog {
            Title = "Wczytaj wersję kalkulatora",
            Filter = "Wersja kalkulatora (*.json)|*.json|Wszystkie pliki (*.*)|*.*",
            DefaultExt = ".json",
            Multiselect = false,
            CheckFileExists = true
        };

        if (dialog.ShowDialog() != true) return;

        await _viewModel.LoadVersionAsync(dialog.FileName);
    }

    private string CreateDefaultVersionFileName() {
        var calculatorName = string.IsNullOrWhiteSpace(_viewModel.CalculatorName)
            ? "wersja_kalkulatora"
            : _viewModel.CalculatorName;

        foreach (var invalidCharacter in Path.GetInvalidFileNameChars())
            calculatorName = calculatorName.Replace(invalidCharacter, '_');

        return $"{calculatorName}_{DateTime.Now:yyyyMMdd_HHmm}.json";
    }

    private void BackButton_Click(object sender, RoutedEventArgs eventArguments) {
        _viewModel.RequestClose();
    }
}