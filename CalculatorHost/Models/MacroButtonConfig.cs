namespace CalculatorHost.Models;

public class MacroButtonConfig {
    public string Label { get; set; } = string.Empty;
    public string MacroName { get; set; } = string.Empty;
    public string Tooltip { get; set; } = string.Empty;
    public bool RefreshLayoutAfterRun { get; set; }
    public bool IsSheetButton { get; set; }
    public bool IsActiveXCommandButton { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string OleObjectName { get; set; } = string.Empty;
    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public int ZIndex { get; set; }
}

public class MacroButtonsFile {
    public List<MacroButtonConfig> Buttons { get; init; } = [];
}