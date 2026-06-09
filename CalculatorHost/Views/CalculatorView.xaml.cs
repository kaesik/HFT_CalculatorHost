using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CalculatorHost.Models;
using CalculatorHost.Rendering;
using CalculatorHost.Services;
using CalculatorHost.ViewModels;
using Microsoft.Win32;

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

    private async void SaveVersionButton_Click(object sender, RoutedEventArgs eventArguments) {
        string defaultFilePath;

        try {
            defaultFilePath = _viewModel.GetDefaultVersionFilePath();
        }
        catch (Exception exception) {
            MessageBox.Show(
                exception.Message,
                "Nie można zapisać wersji",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        var dialog = new SaveFileDialog {
            Title = "Zapisz wersję kalkulatora",
            Filter = "Wersja kalkulatora (*.json)|*.json",
            DefaultExt = ".json",
            AddExtension = true,
            InitialDirectory = Path.GetDirectoryName(defaultFilePath),
            FileName = Path.GetFileName(defaultFilePath)
        };

        if (dialog.ShowDialog() != true)
            return;

        await _viewModel.SaveVersionAsync(dialog.FileName);
    }

    private async void LoadVersionButton_Click(object sender, RoutedEventArgs eventArguments) {
        List<CalculatorVersionFileInfo> matchingVersionFiles;
        string versionsDirectory;

        try {
            matchingVersionFiles = _viewModel.GetMatchingVersionFiles();
            versionsDirectory = _viewModel.GetVersionsDirectory();
        }
        catch (Exception exception) {
            MessageBox.Show(
                exception.Message,
                "Nie można wczytać wersji",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        if (matchingVersionFiles.Count == 0) {
            MessageBox.Show(
                $"Nie znaleziono żadnych wersji pasujących do aktualnie otwartego kalkulatora.\n\nFolder wersji:\n{versionsDirectory}",
                "Brak pasujących wersji",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var selectedVersionFile = ShowVersionSelectionDialog(matchingVersionFiles);

        if (selectedVersionFile == null)
            return;

        await _viewModel.LoadVersionAsync(selectedVersionFile.FilePath);
    }

    private CalculatorVersionFileInfo? ShowVersionSelectionDialog(
        IReadOnlyList<CalculatorVersionFileInfo> matchingVersionFiles) {
        var owner = Window.GetWindow(this);

        var listBox = new ListBox {
            ItemsSource = matchingVersionFiles,
            DisplayMemberPath = nameof(CalculatorVersionFileInfo.DisplayName),
            Margin = new Thickness(12),
            MinHeight = 220
        };

        if (matchingVersionFiles.Count > 0)
            listBox.SelectedIndex = 0;

        var descriptionText = new TextBlock {
            Text = "Wybierz wersję pasującą do aktualnie otwartego kalkulatora:",
            Margin = new Thickness(12, 12, 12, 0),
            TextWrapping = TextWrapping.Wrap
        };

        var loadButton = new Button {
            Content = "Wczytaj",
            Width = 100,
            Height = 32,
            Margin = new Thickness(0, 0, 8, 0),
            IsDefault = true
        };

        var cancelButton = new Button {
            Content = "Anuluj",
            Width = 100,
            Height = 32,
            IsCancel = true
        };

        var buttonsPanel = new StackPanel {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(12)
        };

        buttonsPanel.Children.Add(loadButton);
        buttonsPanel.Children.Add(cancelButton);

        var panel = new DockPanel();
        DockPanel.SetDock(descriptionText, Dock.Top);
        DockPanel.SetDock(buttonsPanel, Dock.Bottom);
        panel.Children.Add(descriptionText);
        panel.Children.Add(buttonsPanel);
        panel.Children.Add(listBox);

        var window = new Window {
            Title = "Wczytaj wersję kalkulatora",
            Owner = owner,
            Width = 620,
            Height = 420,
            MinWidth = 480,
            MinHeight = 300,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Content = panel
        };

        loadButton.Click += (_, _) => {
            if (listBox.SelectedItem == null)
                return;

            window.DialogResult = true;
        };

        listBox.MouseDoubleClick += (_, _) => {
            if (listBox.SelectedItem == null)
                return;

            window.DialogResult = true;
        };

        return window.ShowDialog() == true
            ? listBox.SelectedItem as CalculatorVersionFileInfo
            : null;
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