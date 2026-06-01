namespace CalculatorHost.Models;

public enum MacroButtonActionType {
    Macro,
    ActiveXClick
}

public class MacroButtonConfig {
    public string Label { get; set; } = string.Empty;
    public string MacroName { get; set; } = string.Empty;
    public string Tooltip { get; set; } = string.Empty;
    public bool RefreshLayoutAfterRun { get; set; }

    public MacroButtonActionType ActionType { get; set; } = MacroButtonActionType.Macro;
    public string SheetName { get; set; } = string.Empty;
    public string ShapeName { get; set; } = string.Empty;
    public string OleObjectName { get; set; } = string.Empty;

    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public int ZIndex { get; set; }

    public bool IsSheetButton => Width > 0.0 && Height > 0.0;
}

public class MacroButtonsFile {
    public List<MacroButtonConfig> Buttons { get; init; } = [];
}