using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
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
    internal Border? SelectedCellHighlight { get; set; }
    internal bool IsUpdatingValues { get; set; }

    public Canvas Canvas { get; }
}

public class SheetRenderer {
    private const double MinimumCellSize = 2.0;
    private const double VisibleBorderThickness = 0.75;
    private const int BorderOverlayZIndex = 900;
    private const int SelectedCellHighlightZIndex = 2500;
    private static readonly Color UniformBorderColor = Color.FromRgb(178, 178, 178);
    private static readonly Color SelectedCellBorderColor = Color.FromRgb(37, 99, 235);
    private static Style? _flatComboBoxStyle;

    public static RenderedSheet RenderSheet(
        SheetModel sheet,
        Action<int, int, string> onInputChanged,
        Action<MacroButtonConfig>? onMacroButtonClicked = null,
        Action<CellModel>? onCellSelected = null) {
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
                onInputChanged,
                onCellSelected,
                () => UpdateSelectedCellHighlight(
                    renderedSheet,
                    canvas,
                    cellX,
                    cellY,
                    cellWidth,
                    cellHeight));

            Canvas.SetLeft(element, cellX);
            Canvas.SetTop(element, cellY);
            canvas.Children.Add(element);
        }

        DrawUniformCellGridOverlay(
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
        Action<int, int, string> onInputChanged,
        Action<CellModel>? onCellSelected,
        Action showSelectedCellHighlight) {
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

        RegisterCellSelectionHandlers(
            grid,
            content,
            cell,
            onCellSelected,
            showSelectedCellHighlight);

        return grid;
    }

    private static void RegisterCellSelectionHandlers(
        FrameworkElement cellElement,
        FrameworkElement contentElement,
        CellModel cell,
        Action<CellModel>? onCellSelected,
        Action showSelectedCellHighlight) {
        if (onCellSelected == null)
            return;

        void SelectCell() {
            showSelectedCellHighlight();
            onCellSelected(cell);
        }

        MouseButtonEventHandler mouseHandler = (_, _) => SelectCell();
        KeyboardFocusChangedEventHandler focusHandler = (_, _) => SelectCell();

        cellElement.AddHandler(UIElement.PreviewMouseLeftButtonDownEvent, mouseHandler, true);
        cellElement.AddHandler(UIElement.MouseLeftButtonDownEvent, mouseHandler, true);
        cellElement.AddHandler(UIElement.GotKeyboardFocusEvent, focusHandler, true);

        contentElement.AddHandler(UIElement.PreviewMouseLeftButtonDownEvent, mouseHandler, true);
        contentElement.AddHandler(UIElement.MouseLeftButtonDownEvent, mouseHandler, true);
        contentElement.AddHandler(UIElement.GotKeyboardFocusEvent, focusHandler, true);
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
            // ignored
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
            FontSize = NormalizeFontSize(cell.FontSize),
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

        return new Border {
            Margin = new Thickness(0),
            Background = new SolidColorBrush(cell.BackgroundColor),
            BorderBrush = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(0),
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
            FontSize = NormalizeFontSize(cell.FontSize),
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
            FontSize = NormalizeFontSize(cell.FontSize),
            FontWeight = cell.IsBold ? FontWeights.Bold : FontWeights.Normal,
            Foreground = new SolidColorBrush(cell.ForegroundColor),
            VerticalAlignment = VerticalAlignment.Stretch,
            VerticalContentAlignment = cell.VerticalContentAlignment,
            HorizontalContentAlignment = TextAlignmentToHorizontalAlignment(cell.TextAlignment),
            Padding = new Thickness(4, 0, 4, 0),
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            IsEditable = false,
            Cursor = Cursors.Hand,
            Style = GetFlatComboBoxStyle()
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

            onInputChanged(cell.Row, cell.Column, selectedValue);
        };

        renderedSheet.ComboBoxControls[(cell.Row, cell.Column)] = comboBox;
        return comboBox;
    }


    private static void UpdateSelectedCellHighlight(
        RenderedSheet renderedSheet,
        Canvas canvas,
        double left,
        double top,
        double width,
        double height) {
        var highlight = renderedSheet.SelectedCellHighlight;

        if (highlight == null) {
            highlight = new Border {
                BorderBrush = new SolidColorBrush(SelectedCellBorderColor),
                BorderThickness = new Thickness(2),
                Background = Brushes.Transparent,
                IsHitTestVisible = false,
                SnapsToDevicePixels = true,
                UseLayoutRounding = true
            };

            renderedSheet.SelectedCellHighlight = highlight;
            Panel.SetZIndex(highlight, SelectedCellHighlightZIndex);
            canvas.Children.Add(highlight);
        }

        Canvas.SetLeft(highlight, RoundLayoutPosition(left));
        Canvas.SetTop(highlight, RoundLayoutPosition(top));
        highlight.Width = Math.Max(width, MinimumCellSize);
        highlight.Height = Math.Max(height, MinimumCellSize);
    }

    private static void DrawUniformCellGridOverlay(
        Canvas canvas,
        SheetModel sheet,
        Dictionary<int, double> columnPositions,
        Dictionary<int, double> rowPositions) {
        var gridLines = CollectUniformCellGridLines(sheet);

        foreach (var line in gridLines.Values)
            if (line.IsHorizontal)
                DrawHorizontalBorderLine(canvas, line, columnPositions, rowPositions);
            else
                DrawVerticalBorderLine(canvas, line, columnPositions, rowPositions);
    }

    private static Dictionary<BorderLineKey, BorderLine> CollectUniformCellGridLines(SheetModel sheet) {
        var gridLines = new Dictionary<BorderLineKey, BorderLine>();

        foreach (var cell in sheet.Cells.Where(cell => !cell.IsMergedSlave)) {
            RegisterUniformGridLine(
                gridLines,
                true,
                cell.Row,
                cell.Column,
                cell.Column + cell.ColSpan);

            RegisterUniformGridLine(
                gridLines,
                true,
                cell.Row + cell.RowSpan,
                cell.Column,
                cell.Column + cell.ColSpan);

            RegisterUniformGridLine(
                gridLines,
                false,
                cell.Column,
                cell.Row,
                cell.Row + cell.RowSpan);

            RegisterUniformGridLine(
                gridLines,
                false,
                cell.Column + cell.ColSpan,
                cell.Row,
                cell.Row + cell.RowSpan);
        }

        return gridLines;
    }

    private static void RegisterUniformGridLine(
        Dictionary<BorderLineKey, BorderLine> gridLines,
        bool isHorizontal,
        int boundaryIndex,
        int startIndex,
        int endIndex) {
        if (endIndex <= startIndex)
            return;

        for (var index = startIndex; index < endIndex; index++) {
            var key = new BorderLineKey(isHorizontal, boundaryIndex, index, index + 1);

            if (gridLines.ContainsKey(key))
                continue;

            gridLines[key] = new BorderLine {
                IsHorizontal = isHorizontal,
                BoundaryIndex = boundaryIndex,
                StartIndex = index,
                EndIndex = index + 1,
                Thickness = VisibleBorderThickness,
                Color = UniformBorderColor
            };
        }
    }

    private static void DrawHorizontalBorderLine(
        Canvas canvas,
        BorderLine line,
        Dictionary<int, double> columnPositions,
        Dictionary<int, double> rowPositions) {
        if (!rowPositions.TryGetValue(line.BoundaryIndex, out var y)) return;
        if (!columnPositions.TryGetValue(line.StartIndex, out var startX)) return;
        if (!columnPositions.TryGetValue(line.EndIndex, out var endX)) return;

        startX = RoundLayoutPosition(startX);
        endX = RoundLayoutPosition(endX);
        y = RoundLayoutPosition(y);

        var width = endX - startX;

        if (width <= 0.0)
            return;

        AddCanvasBorderLine(
            canvas,
            startX,
            GetBorderLineStart(y, VisibleBorderThickness, canvas.Height),
            width,
            VisibleBorderThickness,
            line.Color);
    }

    private static void DrawVerticalBorderLine(
        Canvas canvas,
        BorderLine line,
        Dictionary<int, double> columnPositions,
        Dictionary<int, double> rowPositions) {
        if (!columnPositions.TryGetValue(line.BoundaryIndex, out var x)) return;
        if (!rowPositions.TryGetValue(line.StartIndex, out var startY)) return;
        if (!rowPositions.TryGetValue(line.EndIndex, out var endY)) return;

        startY = RoundLayoutPosition(startY);
        endY = RoundLayoutPosition(endY);
        x = RoundLayoutPosition(x);

        var height = endY - startY;

        if (height <= 0.0)
            return;

        AddCanvasBorderLine(
            canvas,
            GetBorderLineStart(x, VisibleBorderThickness, canvas.Width),
            startY,
            VisibleBorderThickness,
            height,
            line.Color);
    }

    private static double GetBorderLineStart(double position, double thickness, double maximumPosition) {
        if (position <= 0.0)
            return 0.0;

        if (position >= maximumPosition)
            return Math.Max(0.0, maximumPosition - thickness);

        return position;
    }

    private static double RoundLayoutPosition(double value) {
        return Math.Round(value, MidpointRounding.AwayFromZero);
    }

    private static double NormalizeFontSize(double fontSize) {
        if (double.IsNaN(fontSize) || fontSize <= 0.0)
            return 11.0;

        return Math.Max(9.0, Math.Min(12.0, fontSize));
    }

    private static Style GetFlatComboBoxStyle() {
        if (_flatComboBoxStyle != null)
            return _flatComboBoxStyle;

        const string styleXaml = @"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type ComboBox}"">
    <Setter Property=""OverridesDefaultStyle"" Value=""True"" />
    <Setter Property=""SnapsToDevicePixels"" Value=""True"" />
    <Setter Property=""Template"">
        <Setter.Value>
            <ControlTemplate TargetType=""{x:Type ComboBox}"">
                <Grid Background=""Transparent"" SnapsToDevicePixels=""True"">
                    <ToggleButton
                        Background=""Transparent""
                        BorderBrush=""Transparent""
                        BorderThickness=""0""
                        ClickMode=""Press""
                        Focusable=""False""
                        IsChecked=""{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"">
                        <Grid Background=""Transparent"">
                            <ContentPresenter
                                Margin=""4,0,18,0""
                                HorizontalAlignment=""{TemplateBinding HorizontalContentAlignment}""
                                VerticalAlignment=""Center""
                                Content=""{TemplateBinding SelectionBoxItem}""
                                ContentStringFormat=""{TemplateBinding SelectionBoxItemStringFormat}""
                                ContentTemplate=""{TemplateBinding SelectionBoxItemTemplate}"" />
                            <TextBlock
                                Margin=""0,0,5,0""
                                HorizontalAlignment=""Right""
                                VerticalAlignment=""Center""
                                FontSize=""9""
                                Foreground=""#666666""
                                IsHitTestVisible=""False""
                                Text=""▾"" />
                        </Grid>
                    </ToggleButton>
                    <Popup
                        x:Name=""PART_Popup""
                        AllowsTransparency=""True""
                        Focusable=""False""
                        IsOpen=""{TemplateBinding IsDropDownOpen}""
                        Placement=""Bottom""
                        PopupAnimation=""Slide"">
                        <Border
                            MinWidth=""{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}""
                            Background=""White""
                            BorderBrush=""#B2B2B2""
                            BorderThickness=""1"">
                            <ScrollViewer MaxHeight=""260"" SnapsToDevicePixels=""True"">
                                <ItemsPresenter />
                            </ScrollViewer>
                        </Border>
                    </Popup>
                </Grid>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
</Style>";

        _flatComboBoxStyle = (Style)XamlReader.Parse(styleXaml);
        return _flatComboBoxStyle;
    }

    private static void AddCanvasBorderLine(
        Canvas canvas,
        double left,
        double top,
        double width,
        double height,
        Color color) {
        if (width <= 0.0 || height <= 0.0)
            return;

        var line = new Border {
            Width = width,
            Height = height,
            Background = new SolidColorBrush(color),
            IsHitTestVisible = false,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };

        Canvas.SetLeft(line, left);
        Canvas.SetTop(line, top);
        Panel.SetZIndex(line, BorderOverlayZIndex);
        canvas.Children.Add(line);
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

    private readonly record struct BorderLineKey(
        bool IsHorizontal,
        int BoundaryIndex,
        int StartIndex,
        int EndIndex);

    private sealed class BorderLine {
        public bool IsHorizontal { get; init; }
        public int BoundaryIndex { get; init; }
        public int StartIndex { get; init; }
        public int EndIndex { get; init; }
        public double Thickness { get; set; }
        public Color Color { get; set; }
    }
}