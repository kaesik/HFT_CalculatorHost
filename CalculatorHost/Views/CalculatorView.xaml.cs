using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using CalculatorHost.Rendering;
using CalculatorHost.ViewModels;

namespace CalculatorHost.Views;

public partial class CalculatorView {
    private readonly CalculatorViewModel _viewModel;

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
            SheetScrollViewer.Content = null;
            return;
        }

        var canvas = SheetRenderer.RenderSheet(
            _viewModel.SheetModel,
            OnInputChanged);

        SheetScrollViewer.Content = canvas;

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

    private void BackButton_Click(object sender, RoutedEventArgs eventArguments) {
        _viewModel.RequestClose();
    }
}