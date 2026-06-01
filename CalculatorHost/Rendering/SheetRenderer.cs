using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using CalculatorHost.Models;

namespace CalculatorHost.Rendering;

public sealed class RenderedSheet {
    public RenderedSheet(Canvas canvas) {
        Canvas = canvas;
    }

    internal Dictionary<(int Row, int Column), TextBlock> ReadOnlyControls { get; } = [];
    internal Dictionary<(int Row, int Column), TextBox> TextBoxControls { get; } = [];
    internal Dictionary<(int Row, int Column), ComboBox> ComboBoxControls { get; } = [];
    internal bool IsUpdatingValues { get; set; }

    public Canvas Canvas { get; }
}

public class SheetRenderer {
    private const double MinimumCellSize = 2.0;

    public static RenderedSheet RenderSheet(SheetModel sheet, Action<int, int, string> onInputChanged,
        Action<MacroButtonConfig>? onMacroButtonClicked = null) {
        var columnPositions = CalculatePositions(
            sheet.FirstColumn,
            sheet.MaxColumn,
            sheet.ColumnWidths,
            sheet.DefaultColumnWidth);

        var rowPositions = CalculatePositions(
            sheet.FirstRow,
            sheet.MaxRow,
            sheet.RowHeights,
            sheet.DefaultRowHeight);

        var cellContentWidth = columnPositions.TryGetValue(sheet.MaxColumn + 1, out var width)
            ? width
            : sheet.DefaultColumnWidth;

        var cellContentHeight = rowPositions.TryGetValue(sheet.MaxRow + 1, out var height)
            ? height
            : sheet.DefaultRowHeight;

        var imageContentWidth = sheet.Images.Count == 0
            ? 0.0
            : sheet.Images.Max(image => image.Left + image.Width);

        var imageContentHeight = sheet.Images.Count == 0
            ? 0.0
            : sheet.Images.Max(image => image.Top + image.Height);

        var macroButtonContentWidth = sheet.MacroButtons.Count == 0
            ? 0.0
            : sheet.MacroButtons.Max(button => button.Left + button.Width);

        var macroButtonContentHeight = sheet.MacroButtons.Count == 0
            ? 0.0
            : sheet.MacroButtons.Max(button => button.Top + button.Height);

        var canvas = new Canvas {
            Width = Math.Max(Math.Max(cellContentWidth, imageContentWidth), macroButtonContentWidth),
            Height = Math.Max(Math.Max(cellContentHeight, imageContentHeight), macroButtonContentHeight),
            Background = Brushes.White,
            SnapsToDevicePixels = true,
            ClipToBounds = true
        };

        var renderedSheet = new RenderedSheet(canvas);

        foreach (var cell in sheet.Cells.Where(cell => !cell.IsMergedSlave)) {
            if (!columnPositions.TryGetValue(cell.Column, out var cellX)) continue;
            if (!rowPositions.TryGetValue(cell.Row, out var cellY)) continue;

            var cellWidth = CalculateSpanSize(
                cell.Column,
                cell.ColSpan,
                sheet.ColumnWidths,
                sheet.DefaultColumnWidth);

            var cellHeight = CalculateSpanSize(
                cell.Row,
                cell.RowSpan,
                sheet.RowHeights,
                sheet.DefaultRowHeight);

            if (cellWidth < MinimumCellSize || cellHeight < MinimumCellSize) continue;

            var element = CreateCellElement(
                renderedSheet,
                cell,
                cellWidth,
                cellHeight,
                onInputChanged);

            Canvas.SetLeft(element, cellX);
            Canvas.SetTop(element, cellY);
            canvas.Children.Add(element);
        }

        foreach (var imageModel in sheet.Images.OrderBy(image => image.ZIndex)) {
            var image = CreateImageElement(imageModel);

            if (image == null)
                continue;

            Canvas.SetLeft(image, imageModel.Left);
            Canvas.SetTop(image, imageModel.Top);
            Panel.SetZIndex(image, 1000 + imageModel.ZIndex);
            canvas.Children.Add(image);
        }

        if (onMacroButtonClicked != null)
            foreach (var macroButton in sheet.MacroButtons.Where(button => button.IsSheetButton)
                         .OrderBy(button => button.ZIndex)) {
                var button = CreateMacroButtonElement(macroButton, onMacroButtonClicked);

                Canvas.SetLeft(button, macroButton.Left);
                Canvas.SetTop(button, macroButton.Top);
                Panel.SetZIndex(button, 2000 + macroButton.ZIndex);
                canvas.Children.Add(button);
            }

        return renderedSheet;
    }

    public static void UpdateCellValues(RenderedSheet renderedSheet, SheetModel sheet) {
        renderedSheet.IsUpdatingValues = true;

        try {
            foreach (var cell in sheet.Cells.Where(cell => !cell.IsMergedSlave)) {
                var coordinate = (cell.Row, cell.Column);

                if (renderedSheet.ReadOnlyControls.TryGetValue(coordinate, out var textBlock)) {
                    if (!string.Equals(textBlock.Text, cell.DisplayText, StringComparison.Ordinal))
                        textBlock.Text = cell.DisplayText;

                    continue;
                }

                if (renderedSheet.TextBoxControls.TryGetValue(coordinate, out var textBox)) {
                    if (!textBox.IsKeyboardFocusWithin &&
                        !string.Equals(textBox.Text, cell.DisplayText, StringComparison.Ordinal))
                        textBox.Text = cell.DisplayText;

                    continue;
                }

                if (!renderedSheet.ComboBoxControls.TryGetValue(coordinate, out var comboBox))
                    continue;

                if (!ReferenceEquals(comboBox.ItemsSource, cell.DropdownValues))
                    comboBox.ItemsSource = cell.DropdownValues;

                var selectedValue = cell.DropdownValues.FirstOrDefault(value =>
                    string.Equals(value, cell.DisplayText, StringComparison.CurrentCulture));

                if (selectedValue != null) {
                    if (!Equals(comboBox.SelectedItem, selectedValue))
                        comboBox.SelectedItem = selectedValue;
                }
                else if (comboBox.SelectedIndex != -1) comboBox.SelectedIndex = -1;
            }
        }
        finally {
            renderedSheet.IsUpdatingValues = false;
        }
    }

    private static Dictionary<int, double> CalculatePositions(
        int first,
        int last,
        Dictionary<int, double> sizes,
        double defaultSize) {
        var positions = new Dictionary<int, double>();
        var position = 0.0;

        for (var index = first; index <= last + 1; index++) {
            positions[index] = position;
            position += sizes.GetValueOrDefault(index, defaultSize);
        }

        return positions;
    }

    private static double CalculateSpanSize(
        int startIndex,
        int span,
        Dictionary<int, double> sizes,
        double defaultSize) {
        var total = 0.0;

        for (var index = startIndex; index < startIndex + span; index++)
            total += sizes.GetValueOrDefault(index, defaultSize);

        return total;
    }

    private static FrameworkElement CreateCellElement(
        RenderedSheet renderedSheet,
        CellModel cell,
        double width,
        double height,
        Action<int, int, string> onInputChanged) {
        var content = cell.IsInput
            ? CreateInputControl(renderedSheet, cell, onInputChanged)
            : CreateReadOnlyControl(renderedSheet, cell);

        return new Border {
            Width = width,
            Height = height,
            Background = new SolidColorBrush(cell.BackgroundColor),
            BorderBrush = new SolidColorBrush(cell.BorderColor),
            BorderThickness = new Thickness(
                cell.BorderLeftThickness,
                cell.BorderTopThickness,
                cell.BorderRightThickness,
                cell.BorderBottomThickness),
            SnapsToDevicePixels = true,
            UseLayoutRounding = true,
            Child = content
        };
    }

    private static Image? CreateImageElement(SheetImageModel imageModel) {
        try {
            using var stream = new MemoryStream(imageModel.ImageBytes, false);

            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.StreamSource = stream;
            bitmap.EndInit();
            bitmap.Freeze();

            return new Image {
                Width = imageModel.Width,
                Height = imageModel.Height,
                Source = bitmap,
                Stretch = Stretch.Fill,
                IsHitTestVisible = false,
                SnapsToDevicePixels = true,
                UseLayoutRounding = true
            };
        }
        catch {
            return null;
        }
    }

    private static Button CreateMacroButtonElement(
        MacroButtonConfig macroButton,
        Action<MacroButtonConfig> onMacroButtonClicked) {
        var fallbackLabel = !string.IsNullOrWhiteSpace(macroButton.ShapeName)
            ? macroButton.ShapeName
            : macroButton.OleObjectName;

        var button = new Button {
            Width = Math.Max(macroButton.Width, MinimumCellSize),
            Height = Math.Max(macroButton.Height, MinimumCellSize),
            Content = string.IsNullOrWhiteSpace(macroButton.Label)
                ? fallbackLabel
                : macroButton.Label,
            ToolTip = string.IsNullOrWhiteSpace(macroButton.Tooltip)
                ? $"Uruchamia makro: {macroButton.MacroName}"
                : macroButton.Tooltip,
            Padding = new Thickness(4, 1, 4, 1),
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center
        };

        button.Click += (_, _) => onMacroButtonClicked(macroButton);

        return button;
    }

    private static FrameworkElement CreateReadOnlyControl(
        RenderedSheet renderedSheet,
        CellModel cell) {
        var textBlock = new TextBlock {
            Text = cell.DisplayText,
            FontSize = Math.Max(cell.FontSize, 7.0),
            FontWeight = cell.IsBold ? FontWeights.Bold : FontWeights.Normal,
            FontStyle = cell.IsItalic ? FontStyles.Italic : FontStyles.Normal,
            Foreground = new SolidColorBrush(cell.ForegroundColor),
            TextAlignment = cell.TextAlignment,
            VerticalAlignment = cell.VerticalContentAlignment,
            TextWrapping = cell.WrapText ? TextWrapping.Wrap : TextWrapping.NoWrap,
            TextTrimming = cell.WrapText ? TextTrimming.None : TextTrimming.CharacterEllipsis,
            Padding = new Thickness(3, 1, 3, 1),
            IsHitTestVisible = false
        };

        renderedSheet.ReadOnlyControls[(cell.Row, cell.Column)] = textBlock;
        return textBlock;
    }

    private static FrameworkElement CreateInputControl(
        RenderedSheet renderedSheet,
        CellModel cell,
        Action<int, int, string> onInputChanged) {
        if (cell.InputType == CellInputType.ComboBox && cell.DropdownValues.Count > 0)
            return CreateComboBoxInput(renderedSheet, cell, onInputChanged);

        return CreateTextBoxInput(renderedSheet, cell, onInputChanged);
    }

    private static TextBox CreateTextBoxInput(
        RenderedSheet renderedSheet,
        CellModel cell,
        Action<int, int, string> onInputChanged) {
        var textBox = new TextBox {
            Text = cell.DisplayText,
            FontSize = Math.Max(cell.FontSize, 7.0),
            FontWeight = cell.IsBold ? FontWeights.Bold : FontWeights.Normal,
            FontStyle = cell.IsItalic ? FontStyles.Italic : FontStyles.Normal,
            Foreground = new SolidColorBrush(cell.ForegroundColor),
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            TextAlignment = cell.TextAlignment,
            VerticalAlignment = VerticalAlignment.Stretch,
            VerticalContentAlignment = cell.VerticalContentAlignment,
            TextWrapping = cell.WrapText ? TextWrapping.Wrap : TextWrapping.NoWrap,
            Padding = new Thickness(3, 1, 3, 1),
            AcceptsReturn = false,
            AcceptsTab = false,
            IsUndoEnabled = true,
            CaretBrush = new SolidColorBrush(
                cell.ForegroundColor == Colors.White ? Colors.Black : cell.ForegroundColor)
        };

        textBox.GotFocus += (_, _) => textBox.SelectAll();

        textBox.LostFocus += (_, _) =>
            onInputChanged(cell.Row, cell.Column, textBox.Text);

        textBox.KeyDown += (_, eventArguments) => {
            switch (eventArguments.Key) {
                case Key.Return:
                    Keyboard.ClearFocus();
                    eventArguments.Handled = true;
                    break;
                case Key.Escape:
                    textBox.Text = cell.DisplayText;
                    Keyboard.ClearFocus();
                    eventArguments.Handled = true;
                    break;
            }
        };

        renderedSheet.TextBoxControls[(cell.Row, cell.Column)] = textBox;
        return textBox;
    }

    private static ComboBox CreateComboBoxInput(
        RenderedSheet renderedSheet,
        CellModel cell,
        Action<int, int, string> onInputChanged) {
        var comboBox = new ComboBox {
            ItemsSource = cell.DropdownValues,
            FontSize = Math.Max(cell.FontSize, 7.0),
            FontWeight = cell.IsBold ? FontWeights.Bold : FontWeights.Normal,
            Foreground = new SolidColorBrush(cell.ForegroundColor),
            VerticalAlignment = VerticalAlignment.Stretch,
            VerticalContentAlignment = cell.VerticalContentAlignment,
            Padding = new Thickness(2, 0, 2, 0),
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            IsEditable = false
        };

        var selectedDisplayValue = cell.DropdownValues.FirstOrDefault(value =>
            string.Equals(value, cell.DisplayText, StringComparison.CurrentCulture));

        if (selectedDisplayValue != null)
            comboBox.SelectedItem = selectedDisplayValue;

        comboBox.SelectionChanged += (_, _) => {
            if (renderedSheet.IsUpdatingValues)
                return;

            if (comboBox.SelectedItem is not string selectedValue)
                return;

            var inputRow = cell.InputTargetRow ?? cell.Row;
            var inputColumn = cell.InputTargetColumn ?? cell.Column;
            var inputValue = cell.DropdownWritesSelectedIndex
                ? (comboBox.SelectedIndex + 1).ToString(CultureInfo.InvariantCulture)
                : selectedValue;

            onInputChanged(inputRow, inputColumn, inputValue);
        };

        renderedSheet.ComboBoxControls[(cell.Row, cell.Column)] = comboBox;
        return comboBox;
    }
}