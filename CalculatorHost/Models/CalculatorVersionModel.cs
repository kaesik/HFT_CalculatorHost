namespace CalculatorHost.Models;

public class CalculatorVersionModel {
    public int FormatVersion { get; init; } = 1;
    public string CalculatorName { get; init; } = string.Empty;
    public string SheetName { get; init; } = string.Empty;
    public DateTime CreatedAt { get; init; } = DateTime.Now;
    public List<CalculatorVersionCellModel> Values { get; init; } = [];
}

public class CalculatorVersionCellModel {
    public int Row { get; init; }
    public int Column { get; init; }
    public string Value { get; init; } = string.Empty;
    public string InputType { get; init; } = string.Empty;
}
