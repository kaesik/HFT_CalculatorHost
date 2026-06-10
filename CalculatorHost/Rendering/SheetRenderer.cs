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
    private const double SoftGridLineThickness = 0.5;
    private const double VisibleBorderMinimumThickness = 1.0;

    public static RenderedSheet RenderSheet(
        SheetModel sheet,
        Action<int, int, string> onInputChanged,
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

        var sheetMacroButtons = sheet.MacroButtons
            .Where(button => button.IsSheetButton)
            .ToList();

        var macroButtonContentWidth = sheetMacroButtons.Count == 0
            ? 0.0
            : sheetMacroButtons.Max(button => button.Left + button.Width);

        var macroButtonContentHeight = sheetMacroButtons.Count == 0
            ? 0.0
            : sheetMacroButtons.Max(button => button.Top + button.Height);

        var canvas = new Canvas {
            Width = Math.Max(Math.Max(cellContentWidth, imageContentWidth), macroButtonContentWidth),
            Height = Math.Max(Math.Max(cellContentHeight, imageContentHeight), macroButtonContentHeight),
            Background = Brushes.White,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true,
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

        DrawExcelBorderOverlay(
            canvas,
            sheet,
            columnPositions,
            rowPositions);

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
            foreach (var macroButton in sheetMacroButtons.OrderBy(button => button.ZIndex)) {
                var button = CreateMacroButtonElement(macroButton, onMacroButtonClicked);

                Canvas.SetLeft(button, macroButton.Left);
                Canvas.SetTop(button, macroButton.Top);
                Panel.SetZIndex(button, 3000 + macroButton.ZIndex);
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
            ? CreateInputControl(renderedSheet, cell, width, height, onInputChanged)
            : CreateReadOnlyControl(renderedSheet, cell);

        var grid = new Grid {
            Width = width,
            Height = height,
            Background = new SolidColorBrush(cell.BackgroundColor),
            ClipToBounds = true,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };

        var baseBorder = new Border {
            Background = new SolidColorBrush(cell.BackgroundColor),
            BorderBrush = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            SnapsToDevicePixels = true,
            UseLayoutRounding = true,
            Child = content
        };

        grid.Children.Add(baseBorder);
        AddExcelBorderLines(grid, cell);

        return grid;
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
            Padding = new Thickness(5, 1, 5, 1),
            FontSize = 11,
            FontWeight = FontWeights.SemiBold,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center,
            Cursor = Cursors.Hand,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };

        try {
            if (Application.Current.TryFindResource("MacroButton") is Style style)
                button.Style = style;
        }
        catch {
        }

        button.Click += (_, _) => onMacroButtonClicked(macroButton);

        return button;
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
            Padding = new Thickness(4, 1, 4, 1),
            IsHitTestVisible = false,
            SnapsToDevicePixels = true
        };

        TextOptions.SetTextFormattingMode(textBlock, TextFormattingMode.Display);
        TextOptions.SetTextRenderingMode(textBlock, TextRenderingMode.ClearType);

        renderedSheet.ReadOnlyControls[(cell.Row, cell.Column)] = textBlock;
        return textBlock;
    }

    private static FrameworkElement CreateInputControl(
        RenderedSheet renderedSheet,
        CellModel cell,
        double width,
        double height,
        Action<int, int, string> onInputChanged) {
        FrameworkElement control = cell.InputType == CellInputType.ComboBox && cell.DropdownValues.Count > 0
            ? CreateComboBoxInput(renderedSheet, cell, onInputChanged)
            : CreateTextBoxInput(renderedSheet, cell, onInputChanged);

        var isSmallCell = width < 26.0 || height < 18.0;

        return new Border {
            Margin = isSmallCell ? new Thickness(0) : new Thickness(1.5),
            Background = CreateInputBackgroundBrush(cell.BackgroundColor),
            BorderBrush = new SolidColorBrush(GetInputBorderColor(cell.InputType)),
            BorderThickness = isSmallCell ? new Thickness(0) : new Thickness(1),
            CornerRadius = isSmallCell ? new CornerRadius(0) : new CornerRadius(3),
            SnapsToDevicePixels = true,
            UseLayoutRounding = true,
            Child = control
        };
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
            Padding = new Thickness(4, 0, 4, 0),
            AcceptsReturn = false,
            AcceptsTab = false,
            IsUndoEnabled = true,
            Cursor = Cursors.IBeam,
            CaretBrush = new SolidColorBrush(
                cell.ForegroundColor == Colors.White ? Colors.Black : cell.ForegroundColor)
        };

        TextOptions.SetTextFormattingMode(textBox, TextFormattingMode.Display);
        TextOptions.SetTextRenderingMode(textBox, TextRenderingMode.ClearType);

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
            HorizontalContentAlignment = TextAlignmentToHorizontalAlignment(cell.TextAlignment),
            Padding = new Thickness(4, 0, 4, 0),
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            IsEditable = false,
            Cursor = Cursors.Hand
        };

        TextOptions.SetTextFormattingMode(comboBox, TextFormattingMode.Display);
        TextOptions.SetTextRenderingMode(comboBox, TextRenderingMode.ClearType);

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


    private static void DrawExcelBorderOverlay(
        Canvas canvas,
        SheetModel sheet,
        Dictionary<int, double> columnPositions,
        Dictionary<int, double> rowPositions) {
        foreach (var cell in sheet.Cells) {
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

            AddCanvasBorderLine(
                canvas,
                cellX,
                cellY,
                cellWidth,
                Math.Max(cell.BorderTopThickness, 0.0),
                cell.BorderColor,
                900);

            AddCanvasBorderLine(
                canvas,
                cellX,
                cellY + cellHeight - Math.Max(cell.BorderBottomThickness, 0.0),
                cellWidth,
                Math.Max(cell.BorderBottomThickness, 0.0),
                cell.BorderColor,
                900);

            AddCanvasBorderLine(
                canvas,
                cellX,
                cellY,
                Math.Max(cell.BorderLeftThickness, 0.0),
                cellHeight,
                cell.BorderColor,
                900);

            AddCanvasBorderLine(
                canvas,
                cellX + cellWidth - Math.Max(cell.BorderRightThickness, 0.0),
                cellY,
                Math.Max(cell.BorderRightThickness, 0.0),
                cellHeight,
                cell.BorderColor,
                900);
        }
    }

    private static void AddCanvasBorderLine(
        Canvas canvas,
        double left,
        double top,
        double width,
        double height,
        Color color,
        int zIndex) {
        if (width <= 0.0 || height <= 0.0)
            return;

        var line = new Border {
            Width = Math.Max(width, VisibleBorderMinimumThickness),
            Height = Math.Max(height, VisibleBorderMinimumThickness),
            Background = new SolidColorBrush(color),
            IsHitTestVisible = false,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };

        Canvas.SetLeft(line, left);
        Canvas.SetTop(line, top);
        Panel.SetZIndex(line, zIndex);
        canvas.Children.Add(line);
    }

    private static void AddExcelBorderLines(Grid grid, CellModel cell) {
        AddBorderLine(
            grid,
            cell.BorderTopThickness,
            cell.BorderColor,
            VerticalAlignment.Top,
            HorizontalAlignment.Stretch,
            null,
            cell.BorderTopThickness);

        AddBorderLine(
            grid,
            cell.BorderBottomThickness,
            cell.BorderColor,
            VerticalAlignment.Bottom,
            HorizontalAlignment.Stretch,
            null,
            cell.BorderBottomThickness);

        AddBorderLine(
            grid,
            cell.BorderLeftThickness,
            cell.BorderColor,
            VerticalAlignment.Stretch,
            HorizontalAlignment.Left,
            cell.BorderLeftThickness,
            null);

        AddBorderLine(
            grid,
            cell.BorderRightThickness,
            cell.BorderColor,
            VerticalAlignment.Stretch,
            HorizontalAlignment.Right,
            cell.BorderRightThickness,
            null);
    }

    private static void AddBorderLine(
        Grid grid,
        double thickness,
        Color color,
        VerticalAlignment verticalAlignment,
        HorizontalAlignment horizontalAlignment,
        double? width,
        double? height) {
        if (thickness <= 0.0)
            return;

        var line = new Border {
            Background = new SolidColorBrush(color),
            VerticalAlignment = verticalAlignment,
            HorizontalAlignment = horizontalAlignment,
            IsHitTestVisible = false,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };

        if (width.HasValue)
            line.Width = Math.Max(width.Value, SoftGridLineThickness);

        if (height.HasValue)
            line.Height = Math.Max(height.Value, SoftGridLineThickness);

        Panel.SetZIndex(line, 20);
        grid.Children.Add(line);
    }

    private static Brush CreateSoftGridBrush(Color backgroundColor) {
        var luminance = GetLuminance(backgroundColor);

        var color = luminance > 170.0
            ? Color.FromRgb(226, 232, 240)
            : Color.FromArgb(72, 255, 255, 255);

        return new SolidColorBrush(color);
    }

    private static Brush CreateInputBackgroundBrush(Color backgroundColor) {
        var luminance = GetLuminance(backgroundColor);

        var color = luminance > 245.0
            ? Color.FromRgb(255, 255, 255)
            : Color.FromArgb(86, 255, 255, 255);

        return new SolidColorBrush(color);
    }

    private static Color GetInputBorderColor(CellInputType inputType) {
        return inputType == CellInputType.ComboBox
            ? Color.FromRgb(37, 99, 235)
            : Color.FromRgb(22, 163, 74);
    }

    private static double GetLuminance(Color color) {
        return 0.299 * color.R + 0.587 * color.G + 0.114 * color.B;
    }

    private static HorizontalAlignment TextAlignmentToHorizontalAlignment(TextAlignment textAlignment) {
        return textAlignment switch {
            TextAlignment.Center => HorizontalAlignment.Center,
            TextAlignment.Right => HorizontalAlignment.Right,
            TextAlignment.Justify => HorizontalAlignment.Stretch,
            _ => HorizontalAlignment.Left
        };
    }
}