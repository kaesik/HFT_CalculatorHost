using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using CalculatorHost.Models;
using CalculatorHost.Rendering;
using CalculatorHost.ViewModels;

namespace CalculatorHost.Views;

public partial class CalculatorView {
    private readonly CalculatorViewModel _viewModel;
    private RenderedSheet? _renderedSheet;
    private SheetModel? _renderedSheetModel;

    public CalculatorView(CalculatorViewModel viewModel) {
        InitializeComponent();

        _viewModel = viewModel;
        DataContext = viewModel;

        viewModel.PropertyChanged += OnViewModelPropertyChanged;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs eventArguments) {
        switch (eventArguments.PropertyName) {
            case nameof(CalculatorViewModel.SheetModel):
                UpdateOrRenderSheet();
                break;
            case nameof(CalculatorViewModel.MacroButtons):
                RebuildMacroButtons();
                break;
        }
    }

    private void UpdateOrRenderSheet() {
        var sheetModel = _viewModel.SheetModel;

        if (sheetModel == null) {
            SheetScrollViewer.Content = null;
            _renderedSheet = null;
            _renderedSheetModel = null;
            return;
        }

        var renderingStopwatch = Stopwatch.StartNew();

        if (_renderedSheet != null && ReferenceEquals(_renderedSheetModel, sheetModel)) {
            SheetRenderer.UpdateCellValues(_renderedSheet, sheetModel);
            renderingStopwatch.Stop();
            _viewModel.ReportRenderingDuration(renderingStopwatch.Elapsed);
            return;
        }

        _renderedSheet = SheetRenderer.RenderSheet(
            sheetModel,
            OnInputChanged,
            OnMacroButtonClicked);

        _renderedSheetModel = sheetModel;
        SheetScrollViewer.Content = _renderedSheet.Canvas;
        UpdateSheetInformation(sheetModel);
        renderingStopwatch.Stop();
        _viewModel.ReportRenderingDuration(renderingStopwatch.Elapsed);
    }

    private void UpdateSheetInformation(SheetModel sheetModel) {
        var inputCount = sheetModel.Cells.Count(cell => cell.IsInput);
        var totalCells = sheetModel.Cells.Count(cell => !cell.IsMergedSlave);

        SheetInfoText.Text =
            $"Arkusz: {sheetModel.SheetName}  ·  " +
            $"Wiersze: {sheetModel.MaxRow}  ·  " +
            $"Kolumny: {sheetModel.MaxColumn}  ·  " +
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

    private void OnMacroButtonClicked(MacroButtonConfig macroButton) {
        if (_viewModel.RunMacroCommand.CanExecute(macroButton))
            _viewModel.RunMacroCommand.Execute(macroButton);
    }

    private void OnInputChanged(int row, int column, string value) {
        _viewModel.SetPendingCellValue(row, column, value);
    }

    private void BackButton_Click(object sender, RoutedEventArgs eventArguments) {
        _viewModel.RequestClose();
    }
}