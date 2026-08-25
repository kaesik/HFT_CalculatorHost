using System.Text.Json.Serialization;

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
    public string WorksheetName { get; set; } = string.Empty;
    public string WorksheetCodeName { get; set; } = string.Empty;

    [JsonIgnore] public string VbaModuleName { get; set; } = string.Empty;

    [JsonIgnore] public string VbaProcedureName { get; set; } = string.Empty;

    [JsonIgnore] public string VbaProcedureCode { get; set; } = string.Empty;

    [JsonIgnore] public string VbaReadError { get; set; } = string.Empty;

    [JsonIgnore] public bool HasVbaProcedureCode => !string.IsNullOrWhiteSpace(VbaProcedureCode);

    public double Left { get; set; }
    public double Top { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public int ZIndex { get; set; }
}

public class MacroButtonsFile {
    public List<MacroButtonConfig> Buttons { get; init; } = [];
}