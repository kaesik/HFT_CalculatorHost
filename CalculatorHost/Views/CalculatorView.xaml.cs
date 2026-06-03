using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CalculatorHost.Models;
using CalculatorHost.Rendering;
using CalculatorHost.ViewModels;

namespace CalculatorHost.Views;

public partial class CalculatorView {
    private const double MinimumZoom = 0.75;
    private const double MaximumZoom = 1.50;
    private const double ZoomStep = 0.25;

    private readonly CalculatorViewModel _viewModel;
    private bool _isChangingZoomInternally;
    private RenderedSheet? _renderedSheet;
    private bool _renderedSheetContainsNonCellElements;
    private SheetModel? _renderedSheetModel;
    private double _zoomFactor = 1.0;

    public CalculatorView(CalculatorViewModel viewModel) {
        InitializeComponent();

        _viewModel = viewModel;
        DataContext = viewModel;

        ApplyZoom();
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
            SheetContentPresenter.Content = null;
            SheetNameText.Text = "Arkusz: —";
            SheetCounterText.Text = "Komórki: —";
            SheetInfoText.Text = string.Empty;
            _renderedSheet = null;
            _renderedSheetModel = null;
            _renderedSheetContainsNonCellElements = false;
            return;
        }

        var renderingStopwatch = Stopwatch.StartNew();
        var containsNonCellElements = ContainsNonCellElements(sheetModel);

        if (_renderedSheet != null &&
            ReferenceEquals(_renderedSheetModel, sheetModel) &&
            !containsNonCellElements &&
            !_renderedSheetContainsNonCellElements) {
            SheetRenderer.UpdateCellValues(_renderedSheet, sheetModel);
            UpdateSheetInformation(sheetModel);
            renderingStopwatch.Stop();
            _viewModel.ReportRenderingDuration(renderingStopwatch.Elapsed);
            return;
        }

        _renderedSheet = SheetRenderer.RenderSheet(
            sheetModel,
            OnInputChanged,
            OnMacroButtonClicked);

        _renderedSheetModel = sheetModel;
        _renderedSheetContainsNonCellElements = containsNonCellElements;
        SheetContentPresenter.Content = _renderedSheet.Canvas;
        ApplyZoom();
        UpdateSheetInformation(sheetModel);
        renderingStopwatch.Stop();
        _viewModel.ReportRenderingDuration(renderingStopwatch.Elapsed);
    }

    private static bool ContainsNonCellElements(SheetModel sheetModel) {
        return sheetModel.Images.Count > 0 ||
               sheetModel.MacroButtons.Any(button => button.IsSheetButton);
    }

    private void UpdateSheetInformation(SheetModel sheetModel) {
        var inputCount = sheetModel.Cells.Count(cell => cell.IsInput);
        var dropdownCount = sheetModel.Cells.Count(cell => cell.InputType == CellInputType.ComboBox);
        var totalCells = sheetModel.Cells.Count(cell => !cell.IsMergedSlave);
        var rowCount = sheetModel.MaxRow - sheetModel.FirstRow + 1;
        var columnCount = sheetModel.MaxColumn - sheetModel.FirstColumn + 1;

        SheetNameText.Text = $"Arkusz: {sheetModel.SheetName}";
        SheetCounterText.Text = $"{rowCount} w. × {columnCount} kol. · {inputCount} pól · {dropdownCount} list";
        SheetInfoText.Text =
            $"Komórek renderowanych: {totalCells} · " +
            $"Pól edytowalnych: {inputCount} · " +
            $"Dropdownów: {dropdownCount} · " +
            $"Obrazów: {sheetModel.Images.Count}";
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

    private void ZoomSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> eventArguments) {
        if (_isChangingZoomInternally) return;

        _zoomFactor = ClampZoom(eventArguments.NewValue / 100.0);
        ApplyZoom();
    }

    private void ZoomOut_Click(object sender, RoutedEventArgs eventArguments) {
        SetZoom(_zoomFactor - ZoomStep);
    }

    private void ZoomIn_Click(object sender, RoutedEventArgs eventArguments) {
        SetZoom(_zoomFactor + ZoomStep);
    }

    private void ResetZoom_Click(object sender, RoutedEventArgs eventArguments) {
        SetZoom(1.0);
    }

    private void SetZoom(double zoomFactor) {
        _zoomFactor = ClampZoom(zoomFactor);

        _isChangingZoomInternally = true;
        try {
            ZoomSlider.Value = _zoomFactor * 100.0;
        }
        finally {
            _isChangingZoomInternally = false;
        }

        ApplyZoom();
    }

    private void ApplyZoom() {
        if (SheetViewport == null || ZoomValueButton == null)
            return;

        SheetViewport.LayoutTransform = new ScaleTransform(_zoomFactor, _zoomFactor);
        ZoomValueButton.Content = $"{_zoomFactor * 100.0:0}%";
    }

    private static double ClampZoom(double zoomFactor) {
        return Math.Max(MinimumZoom, Math.Min(MaximumZoom, zoomFactor));
    }
}