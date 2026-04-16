using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using CalculatorHost.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalculatorHost.Services;

/// <summary>
///     Reads the first worksheet from a workbook and converts it into a SheetModel
///     that can be rendered by WPF without any further Excel dependency.
///     Must be called from the ExcelWorker STA thread.
/// </summary>
public class SheetReaderService {
    private const double PointsToDips = 96.0 / 72.0;

    public SheetModel ReadFirstSheet(ExcelSessionService session) {
        Excel.Worksheet? worksheet = null;
        Excel.Range? usedRange = null;

        try {
            worksheet = session.GetFirstWorksheet();
            usedRange = worksheet.UsedRange;

            var model = new SheetModel {
                SheetName = worksheet.Name,
                FirstRow = usedRange.Row,
                FirstColumn = usedRange.Column,
                MaxRow = usedRange.Row + usedRange.Rows.Count - 1,
                MaxColumn = usedRange.Column + usedRange.Columns.Count - 1
            };

            ReadColumnWidths(worksheet, model);
            ReadRowHeights(worksheet, model);
            ReadCells(worksheet, usedRange, model);

            return model;
        }
        finally {
            if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            if (worksheet != null) Marshal.ReleaseComObject(worksheet);
        }
    }

    private static void ReadColumnWidths(Excel.Worksheet worksheet, SheetModel model) {
        for (var col = model.FirstColumn; col <= model.MaxColumn; col++) {
            Excel.Range? colRange = null;
            try {
                colRange = (Excel.Range)worksheet.Columns[col];
                var isHidden = (bool)colRange.Hidden;
                var widthDips = isHidden ? 0.0 : Math.Max((double)colRange.Width * PointsToDips, 0.0);
                model.ColumnWidths[col] = widthDips;
            }
            catch {
                model.ColumnWidths[col] = model.DefaultColumnWidth;
            }
            finally {
                if (colRange != null) Marshal.ReleaseComObject(colRange);
            }
        }
    }

    private static void ReadRowHeights(Excel.Worksheet worksheet, SheetModel model) {
        for (var row = model.FirstRow; row <= model.MaxRow; row++) {
            Excel.Range? rowRange = null;
            try {
                rowRange = (Excel.Range)worksheet.Rows[row];
                var isHidden = (bool)rowRange.Hidden;
                var heightDips = isHidden ? 0.0 : Math.Max((double)rowRange.Height * PointsToDips, 0.0);
                model.RowHeights[row] = heightDips;
            }
            catch {
                model.RowHeights[row] = model.DefaultRowHeight;
            }
            finally {
                if (rowRange != null) Marshal.ReleaseComObject(rowRange);
            }
        }
    }

    private void ReadCells(Excel.Worksheet worksheet, Excel.Range usedRange, SheetModel model) {
        var mergedSlaves = new HashSet<(int Row, int Col)>();

        for (var row = model.FirstRow; row <= model.MaxRow; row++) {
            if (model.RowHeights.TryGetValue(row, out var rowH) && rowH == 0.0) continue;

            for (var col = model.FirstColumn; col <= model.MaxColumn; col++) {
                if (model.ColumnWidths.TryGetValue(col, out var colW) && colW == 0.0) continue;

                if (mergedSlaves.Contains((row, col))) {
                    model.Cells.Add(new CellModel { Row = row, Column = col, IsMergedSlave = true });
                    continue;
                }

                Excel.Range? cell = null;
                try {
                    cell = (Excel.Range)worksheet.Cells[row, col];
                    var cellModel = ReadSingleCell(cell, row, col, worksheet);

                    // If this is the top-left of a merge, register all other cells as slaves
                    if (cellModel.RowSpan > 1 || cellModel.ColSpan > 1)
                        for (var mr = row; mr < row + cellModel.RowSpan; mr++)
                        for (var mc = col; mc < col + cellModel.ColSpan; mc++)
                            if (mr != row || mc != col)
                                mergedSlaves.Add((mr, mc));

                    model.Cells.Add(cellModel);
                }
                catch {
                    model.Cells.Add(new CellModel {
                        Row = row,
                        Column = col,
                        DisplayText = string.Empty,
                        BackgroundColor = Colors.White
                    });
                }
                finally {
                    if (cell != null) Marshal.ReleaseComObject(cell);
                }
            }
        }
    }

    private CellModel ReadSingleCell(Excel.Range cell, int row, int col, Excel.Worksheet worksheet) {
        var model = new CellModel { Row = row, Column = col };

        ReadValueAndText(cell, model);
        ReadFont(cell, model);
        ReadBackground(cell, model);
        ReadAlignment(cell, model);
        ReadBorders(cell, model);
        ReadMerge(cell, model);

        if (model.IsInput)
            ReadValidation(cell, model, worksheet);

        return model;
    }

    private static void ReadValueAndText(Excel.Range cell, CellModel model) {
        try {
            model.RawValue = cell.Value2;
            model.DisplayText = cell.Text as string ?? model.RawValue?.ToString() ?? string.Empty;
        }
        catch {
            model.DisplayText = string.Empty;
        }
    }

    private static void ReadFont(Excel.Range cell, CellModel model) {
        Excel.Font? font = null;
        try {
            font = cell.Font;
            model.IsBold = font.Bold is true;
            model.IsItalic = font.Italic is true;

            try {
                model.FontSize = (double)font.Size;
            }
            catch {
                model.FontSize = 11.0;
            }

            try {
                if (font.Color is double colorValue)
                    model.ForegroundColor = OleColorToMediaColor((int)colorValue);
                else
                    model.ForegroundColor = Colors.Black;
            }
            catch {
                model.ForegroundColor = Colors.Black;
            }
        }
        catch {
            // ignored
        }
        finally {
            if (font != null) Marshal.ReleaseComObject(font);
        }
    }

    private static void ReadBackground(Excel.Range cell, CellModel model) {
        Excel.Interior? interior = null;
        try {
            interior = cell.Interior;
            var colorIndex = interior.ColorIndex;

            if (colorIndex is (int)Excel.XlColorIndex.xlColorIndexNone) {
                model.BackgroundColor = Colors.White;
                model.IsInput = false;
                return;
            }

            if (interior.Color is double bgDouble) {
                model.BackgroundColor = OleColorToMediaColor((int)bgDouble);
                model.IsInput = IsColorGreen(model.BackgroundColor);
                model.InputType = model.IsInput ? CellInputType.TextBox : CellInputType.None;
            }
            else
                model.BackgroundColor = Colors.White;
        }
        catch {
            model.BackgroundColor = Colors.White;
        }
        finally {
            if (interior != null) Marshal.ReleaseComObject(interior);
        }
    }

    private static void ReadAlignment(Excel.Range cell, CellModel model) {
        try {
            model.TextAlignment = ExcelHAlignToWpf(cell.HorizontalAlignment);
            model.VerticalContentAlignment = ExcelVAlignToWpf(cell.VerticalAlignment);
            model.WrapText = cell.WrapText is true;
        }
        catch {
            // ignored
        }
    }

    private void ReadBorders(Excel.Range cell, CellModel model) {
        Excel.Borders? borders = null;
        try {
            borders = cell.Borders;
            var dominantColor = Colors.Black;

            model.BorderTopThickness = ReadBorderThickness(borders, Excel.XlBordersIndex.xlEdgeTop, ref dominantColor);
            model.BorderBottomThickness =
                ReadBorderThickness(borders, Excel.XlBordersIndex.xlEdgeBottom, ref dominantColor);
            model.BorderLeftThickness =
                ReadBorderThickness(borders, Excel.XlBordersIndex.xlEdgeLeft, ref dominantColor);
            model.BorderRightThickness =
                ReadBorderThickness(borders, Excel.XlBordersIndex.xlEdgeRight, ref dominantColor);
            model.BorderColor = dominantColor;
        }
        catch {
            // ignored
        }
        finally {
            if (borders != null) Marshal.ReleaseComObject(borders);
        }
    }

    private static double ReadBorderThickness(Excel.Borders borders, Excel.XlBordersIndex borderIndex,
        ref Color dominantColor) {
        Excel.Border? border = null;
        try {
            border = borders[borderIndex];
            var lineStyle = border.LineStyle;
            if (lineStyle is int ls && ls != (int)Excel.XlLineStyle.xlLineStyleNone) {
                var weight = border.Weight;
                if (border.Color is double colorDouble)
                    dominantColor = OleColorToMediaColor((int)colorDouble);
                return ExcelBorderWeightToDips(weight);
            }

            return 0.0;
        }
        catch {
            return 0.0;
        }
        finally {
            if (border != null) Marshal.ReleaseComObject(border);
        }
    }

    private static void ReadMerge(Excel.Range cell, CellModel model) {
        Excel.Range? mergeArea = null;
        try {
            if (cell.MergeCells is not true) return;

            mergeArea = cell.MergeArea;
            model.RowSpan = mergeArea.Rows.Count;
            model.ColSpan = mergeArea.Columns.Count;
        }
        catch {
            // ignored
        }
        finally {
            if (mergeArea != null) Marshal.ReleaseComObject(mergeArea);
        }
    }

    private void ReadValidation(Excel.Range cell, CellModel model, Excel.Worksheet worksheet) {
        Excel.Validation? validation = null;
        try {
            validation = cell.Validation;
            if (validation == null) return;

            var validationType = validation.Type;
            if (validationType != (int)Excel.XlDVType.xlValidateList) return;

            var formula = validation.Formula1;
            if (string.IsNullOrEmpty(formula)) return;

            var values = ParseValidationFormula(formula, worksheet);
            if (values.Count == 0) return;

            model.DropdownValues = values;
            model.InputType = CellInputType.ComboBox;
        }
        catch {
            // ignored
        }
        finally {
            if (validation != null) Marshal.ReleaseComObject(validation);
        }
    }

    private static List<string> ParseValidationFormula(string formula, Excel.Worksheet worksheet) {
        var result = new List<string>();

        if (formula.StartsWith('=')) {
            // Reference to a range
            Excel.Range? range = null;
            try {
                range = worksheet.Range[formula.TrimStart('=')];
                foreach (Excel.Range rangeCell in range.Cells) {
                    var value = rangeCell.Value2?.ToString();
                    if (!string.IsNullOrEmpty(value)) result.Add(value);
                    Marshal.ReleaseComObject(rangeCell);
                }
            }
            catch {
                // ignored
            }
            finally {
                if (range != null) Marshal.ReleaseComObject(range);
            }
        }
        else
            // Comma-separated literal list
            result.AddRange(formula.Split(',').Select(v => v.Trim()).Where(v => !string.IsNullOrEmpty(v)));

        return result;
    }

    private static Color OleColorToMediaColor(int oleColor) {
        // Excel stores colors as BGR (blue in lowest byte)
        var r = (byte)(oleColor & 0xFF);
        var g = (byte)((oleColor >> 8) & 0xFF);
        var b = (byte)((oleColor >> 16) & 0xFF);
        return Color.FromRgb(r, g, b);
    }

    /// <summary>
    ///     Detects whether a color is in the green hue range using HSV color space analysis.
    ///     Covers standard Excel green fills: #00B050, #92D050, #C6EFCE, #00FF00 etc.
    /// </summary>
    private static bool IsColorGreen(Color color) {
        var r = color.R / 255.0;
        var g = color.G / 255.0;
        var b = color.B / 255.0;

        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var delta = max - min;

        if (delta < 0.08) return false; // too achromatic (gray/white/black)
        if (max < 0.15) return false; // too dark

        double hue;
        if (Math.Abs(max - g) < 0.001)
            hue = 60.0 * ((b - r) / delta + 2.0);
        else if (Math.Abs(max - r) < 0.001)
            hue = 60.0 * ((g - b) / delta % 6.0);
        else
            hue = 60.0 * ((r - g) / delta + 4.0);

        if (hue < 0) hue += 360.0;

        var saturation = max == 0 ? 0 : delta / max;

        // Green hue range: approximately 70–165 degrees, saturation > 12%
        return hue is >= 70.0 and <= 165.0 && saturation >= 0.12;
    }

    private static TextAlignment ExcelHAlignToWpf(object hAlign) {
        if (hAlign is not int ha) return TextAlignment.Left;
        return ha switch {
            (int)Excel.XlHAlign.xlHAlignLeft => TextAlignment.Left,
            (int)Excel.XlHAlign.xlHAlignCenter => TextAlignment.Center,
            (int)Excel.XlHAlign.xlHAlignRight => TextAlignment.Right,
            (int)Excel.XlHAlign.xlHAlignJustify => TextAlignment.Justify,
            (int)Excel.XlHAlign.xlHAlignGeneral => TextAlignment.Left,
            _ => TextAlignment.Left
        };
    }

    private static VerticalAlignment ExcelVAlignToWpf(object vAlign) {
        if (vAlign is not int va) return VerticalAlignment.Center;
        return va switch {
            (int)Excel.XlVAlign.xlVAlignTop => VerticalAlignment.Top,
            (int)Excel.XlVAlign.xlVAlignCenter => VerticalAlignment.Center,
            (int)Excel.XlVAlign.xlVAlignBottom => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Center
        };
    }

    private static double ExcelBorderWeightToDips(object weight) {
        if (weight is not int w) return 1.0;
        return w switch {
            (int)Excel.XlBorderWeight.xlHairline => 0.5,
            (int)Excel.XlBorderWeight.xlThin => 1.0,
            (int)Excel.XlBorderWeight.xlMedium => 2.0,
            (int)Excel.XlBorderWeight.xlThick => 3.0,
            _ => 1.0
        };
    }
}