namespace CalculatorHost.Models;

public abstract class MacroButtonConfig {
    public string Label { get; set; } = string.Empty;
    public string MacroName { get; set; } = string.Empty;
    public string Tooltip { get; set; } = string.Empty;
}

public class MacroButtonsFile {
    public List<MacroButtonConfig> Buttons { get; init; } = [];
}