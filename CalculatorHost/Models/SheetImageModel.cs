namespace CalculatorHost.Models;

public class SheetImageModel {
    public string Name { get; init; } = string.Empty;
    public byte[] ImageBytes { get; init; } = [];
    public double Left { get; init; }
    public double Top { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
    public int ZIndex { get; init; }
}