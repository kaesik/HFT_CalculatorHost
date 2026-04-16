using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using CalculatorHost.Models;

namespace CalculatorHost.Rendering;

/// <summary>
///     Converts a SheetModel into a WPF Canvas containing positioned Border/TextBlock/TextBox/ComboBox
///     elements that visually replicate the Excel worksheet.
///     Canvas with absolute positioning is used because the sheet geometry (merged cells, variable
///     column widths, variable row heights) cannot be reliably expressed with Grid rows/columns.
/// </summary>
public class SheetRenderer {
    private const double MinCellSize = 2.0;

    public static Canvas RenderSheet(SheetModel sheet, Action<int, int, string> onInputCommitted) {
        var columnPositions = CalculatePositions(sheet.FirstColumn, sheet.MaxColumn, sheet.ColumnWidths,
            sheet.DefaultColumnWidth);
        var rowPositions = CalculatePositions(sheet.FirstRow, sheet.MaxRow, sheet.RowHeights, sheet.DefaultRowHeight);

        var totalWidth = columnPositions.TryGetValue(sheet.MaxColumn + 1, out var tw) ? tw : sheet.DefaultColumnWidth;
        var totalHeight = rowPositions.TryGetValue(sheet.MaxRow + 1, out var th) ? th : sheet.DefaultRowHeight;

        var canvas = new Canvas {
            Width = totalWidth,
            Height = totalHeight,
            Background = Brushes.White,
            SnapsToDevicePixels = true
        };

        foreach (var cell in sheet.Cells.Where(cell => !cell.IsMergedSlave)) {
            if (!columnPositions.TryGetValue(cell.Column, out var cellX)) continue;
            if (!rowPositions.TryGetValue(cell.Row, out var cellY)) continue;

            var cellWidth = CalculateSpanSize(cell.Column, cell.ColSpan, sheet.ColumnWidths, sheet.DefaultColumnWidth);
            var cellHeight = CalculateSpanSize(cell.Row, cell.RowSpan, sheet.RowHeights, sheet.DefaultRowHeight);

            if (cellWidth < MinCellSize || cellHeight < MinCellSize) continue;

            var element = CreateCellElement(cell, cellWidth, cellHeight, onInputCommitted);

            Canvas.SetLeft(element, cellX);
            Canvas.SetTop(element, cellY);
            canvas.Children.Add(element);
        }

        return canvas;
    }

    private static Dictionary<int, double> CalculatePositions(int first, int last, Dictionary<int, double> sizes,
        double defaultSize) {
        var positions = new Dictionary<int, double>();
        var position = 0.0;

        for (var index = first; index <= last + 1; index++) {
            positions[index] = position;
            var size = sizes.GetValueOrDefault(index, defaultSize);
            position += size;
        }

        return positions;
    }

    private static double CalculateSpanSize(int startIndex, int span, Dictionary<int, double> sizes,
        double defaultSize) {
        var total = 0.0;
        for (var i = startIndex; i < startIndex + span; i++)
            total += sizes.GetValueOrDefault(i, defaultSize);
        return total;
    }

    private static FrameworkElement CreateCellElement(CellModel cell, double width, double height,
        Action<int, int, string> onInputCommitted) {
        var border = new Border {
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
            Child = cell.IsInput ? CreateInputControl(cell, onInputCommitted) : CreateReadOnlyControl(cell)
        };

        return border;
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

    private static FrameworkElement CreateInputControl(CellModel cell, Action<int, int, string> onInputCommitted) {
        if (cell.InputType == CellInputType.ComboBox && cell.DropdownValues.Count > 0)
            return CreateComboBoxInput(cell, onInputCommitted);

        return CreateTextBoxInput(cell, onInputCommitted);
    }

    private static TextBox CreateTextBoxInput(CellModel cell, Action<int, int, string> onInputCommitted) {
        var capturedRow = cell.Row;
        var capturedColumn = cell.Column;

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
            // Set caret brush explicitly for visibility
            CaretBrush = new SolidColorBrush(
                cell.ForegroundColor == Colors.White ? Colors.Black : cell.ForegroundColor)
        };

        textBox.GotFocus += (_, _) => textBox.SelectAll();

        textBox.LostFocus += (_, _) =>
            onInputCommitted(capturedRow, capturedColumn, textBox.Text);

        textBox.KeyDown += (_, e) => {
            switch (e.Key) {
                case Key.Return:
                    onInputCommitted(capturedRow, capturedColumn, textBox.Text);
                    Keyboard.ClearFocus();
                    e.Handled = true;
                    break;
                case Key.Tab:
                    onInputCommitted(capturedRow, capturedColumn, textBox.Text);
                    break;
                case Key.Escape:
                    Keyboard.ClearFocus();
                    e.Handled = true;
                    break;
            }
        };

        return textBox;
    }

    private static ComboBox CreateComboBoxInput(CellModel cell, Action<int, int, string> onInputCommitted) {
        var capturedRow = cell.Row;
        var capturedColumn = cell.Column;

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

        // Set current selection without triggering the event
        var currentText = cell.DisplayText;
        if (cell.DropdownValues.Contains(currentText))
            comboBox.SelectedItem = currentText;

        comboBox.SelectionChanged += (_, _) => {
            if (comboBox.SelectedItem is string selected)
                onInputCommitted(capturedRow, capturedColumn, selected);
        };

        return comboBox;
    }
}