using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using CalculatorHost.Models;

namespace CalculatorHost.Rendering;

public sealed class RenderedSheet {
    public Canvas Canvas { get; init; } = new();
    public Dictionary<(int Row, int Column), FrameworkElement> CellElements { get; } = [];
}

public class SheetRenderer {
    private const double MinimumCellSize = 2.0;

    public static RenderedSheet RenderSheet(SheetModel sheet, Action<int, int, string> onInputChanged) {
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

        var totalWidth = columnPositions.TryGetValue(sheet.MaxColumn + 1, out var width)
            ? width
            : sheet.DefaultColumnWidth;

        var totalHeight = rowPositions.TryGetValue(sheet.MaxRow + 1, out var height)
            ? height
            : sheet.DefaultRowHeight;

        var renderedSheet = new RenderedSheet {
            Canvas = new Canvas {
                Width = totalWidth,
                Height = totalHeight,
                Background = Brushes.White,
                SnapsToDevicePixels = true
            }
        };

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

            var element = CreateCellElement(cell, cellWidth, cellHeight, onInputChanged);

            Canvas.SetLeft(element, cellX);
            Canvas.SetTop(element, cellY);
            renderedSheet.Canvas.Children.Add(element);
            renderedSheet.CellElements[(cell.Row, cell.Column)] = element;
        }

        return renderedSheet;
    }

    public static void UpdateCellValues(RenderedSheet renderedSheet, SheetModel sheet) {
        foreach (var cell in sheet.Cells.Where(cell => !cell.IsMergedSlave)) {
            if (!renderedSheet.CellElements.TryGetValue((cell.Row, cell.Column), out var element))
                continue;

            if (element is not Border border)
                continue;

            switch (border.Child) {
                case TextBlock textBlock:
                    if (textBlock.Text != cell.DisplayText)
                        textBlock.Text = cell.DisplayText;
                    break;
                case TextBox textBox:
                    if (textBox.Text != cell.DisplayText)
                        textBox.Text = cell.DisplayText;
                    break;
                case ComboBox comboBox:
                    UpdateComboBoxValue(comboBox, cell.DisplayText);
                    break;
            }
        }
    }

    private static void UpdateComboBoxValue(ComboBox comboBox, string value) {
        if (Equals(comboBox.SelectedItem, value))
            return;

        comboBox.Tag = true;

        try {
            comboBox.SelectedItem = value;
        }
        finally {
            comboBox.Tag = null;
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
        CellModel cell,
        double width,
        double height,
        Action<int, int, string> onInputChanged) {
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
            Child = cell.IsInput
                ? CreateInputControl(cell, onInputChanged)
                : CreateReadOnlyControl(cell)
        };
    }

    private static FrameworkElement CreateReadOnlyControl(CellModel cell) {
        return new TextBlock {
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
    }

    private static FrameworkElement CreateInputControl(
        CellModel cell,
        Action<int, int, string> onInputChanged) {
        if (cell.InputType == CellInputType.ComboBox && cell.DropdownValues.Count > 0)
            return CreateComboBoxInput(cell, onInputChanged);

        return CreateTextBoxInput(cell, onInputChanged);
    }

    private static TextBox CreateTextBoxInput(
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

        return textBox;
    }

    private static ComboBox CreateComboBoxInput(
        CellModel cell,
        Action<int, int, string> onInputChanged) {
        var comboBox = new ComboBox {
            ItemsSource = cell.DropdownValues,
            FontSize = Math.Max(cell.FontSize, 7.0),
            FontWeight = cell.IsBold ? FontWeights.Bold : FontWeights.Normal,
            VerticalAlignment = VerticalAlignment.Stretch,
            VerticalContentAlignment = cell.VerticalContentAlignment,
            Padding = new Thickness(2, 0, 2, 0),
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            IsEditable = false
        };

        if (cell.DropdownValues.Contains(cell.DisplayText))
            comboBox.SelectedItem = cell.DisplayText;

        comboBox.SelectionChanged += (_, _) => {
            if (comboBox.Tag is true)
                return;

            if (comboBox.SelectedItem is string selectedValue)
                onInputChanged(cell.Row, cell.Column, selectedValue);
        };

        return comboBox;
    }
}