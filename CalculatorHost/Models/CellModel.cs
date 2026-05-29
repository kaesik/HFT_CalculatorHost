using System.Windows;
using System.Windows.Media;

namespace CalculatorHost.Models;

public enum CellInputType {
    None,
    TextBox,
    ComboBox
}

public class CellModel {
    public int Row { get; init; }
    public int Column { get; init; }
    public int RowSpan { get; set; } = 1;
    public int ColSpan { get; set; } = 1;
    public bool IsMergedSlave { get; init; }

    public string DisplayText { get; set; } = string.Empty;
    public object? RawValue { get; set; }

    public Color BackgroundColor { get; set; } = Colors.White;
    public Color ForegroundColor { get; set; } = Colors.Black;

    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public double FontSize { get; set; } = 11.0;

    public TextAlignment TextAlignment { get; set; } = TextAlignment.Left;
    public VerticalAlignment VerticalContentAlignment { get; set; } = VerticalAlignment.Center;

    public bool WrapText { get; set; }

    public double BorderTopThickness { get; set; }
    public double BorderBottomThickness { get; set; }
    public double BorderLeftThickness { get; set; }
    public double BorderRightThickness { get; set; }
    public Color BorderColor { get; set; } = Colors.Black;

    public bool IsInput { get; set; }
    public CellInputType InputType { get; set; } = CellInputType.None;
    public List<string> DropdownValues { get; set; } = [];
    public int? InputTargetRow { get; set; }
    public int? InputTargetColumn { get; set; }
    public bool DropdownWritesSelectedIndex { get; set; }
}