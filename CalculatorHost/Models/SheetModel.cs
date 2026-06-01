namespace CalculatorHost.Models;

public class SheetModel {
    public string SheetName { get; init; } = string.Empty;
    public List<CellModel> Cells { get; set; } = [];
    public List<SheetImageModel> Images { get; set; } = [];
    public List<MacroButtonConfig> MacroButtons { get; set; } = [];
    public Dictionary<int, double> ColumnWidths { get; set; } = new();
    public Dictionary<int, double> RowHeights { get; set; } = new();
    public int FirstRow { get; init; } = 1;
    public int FirstColumn { get; init; } = 1;
    public int MaxRow { get; init; }
    public int MaxColumn { get; init; }
    public double DefaultColumnWidth { get; set; } = 64.0;
    public double DefaultRowHeight { get; set; } = 20.0;

    public double TotalWidth => ColumnWidths.Values.Sum();
    public double TotalHeight => RowHeights.Values.Sum();
}