using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public class SheetReaderService {
    private const double PointsToDips = 96.0 / 72.0;

    private const int MaximumDirectReadCellCount = 3000;
    private const int AdditionalRowsAfterLastContent = 6;
    private const int AdditionalColumnsAfterLastContent = 4;

    private const int ExcelColorIndexNone = -4142;

    private const int ExcelBorderEdgeLeft = 7;
    private const int ExcelBorderEdgeTop = 8;
    private const int ExcelBorderEdgeBottom = 9;
    private const int ExcelBorderEdgeRight = 10;

    private const int ExcelLineStyleNone = -4142;
    private const int ExcelValidationTypeList = 3;

    private const int ExcelFindLookInFormulas = -4123;
    private const int ExcelSearchOrderByRows = 1;
    private const int ExcelSearchOrderByColumns = 2;
    private const int ExcelSearchDirectionPrevious = 2;

    private const int ExcelHorizontalAlignmentGeneral = 1;
    private const int ExcelHorizontalAlignmentLeft = -4131;
    private const int ExcelHorizontalAlignmentCenter = -4108;
    private const int ExcelHorizontalAlignmentRight = -4152;
    private const int ExcelHorizontalAlignmentJustify = -4130;

    private const int ExcelVerticalAlignmentTop = -4160;
    private const int ExcelVerticalAlignmentCenter = -4108;
    private const int ExcelVerticalAlignmentBottom = -4107;

    private const int ExcelBorderWeightHairline = 1;
    private const int ExcelBorderWeightThin = 2;
    private const int ExcelBorderWeightMedium = -4138;
    private const int ExcelBorderWeightThick = 4;

    private const string InputColorLightGreen = "#92D050";
    private const string InputColorGreen = "#00B050";
    private const string DropdownColor = "#00B0F0";

    public SheetModel ReadFirstSheet(ExcelSessionService session) {
        dynamic? worksheet = null;
        dynamic? usedRange = null;
        dynamic? usedRows = null;
        dynamic? usedColumns = null;

        try {
            worksheet = session.GetFirstWorksheet();
            usedRange = worksheet.UsedRange;
            usedRows = usedRange.Rows;
            usedColumns = usedRange.Columns;

            var usedFirstRow = Convert.ToInt32(usedRange.Row);
            var usedFirstColumn = Convert.ToInt32(usedRange.Column);
            var usedMaximumRow = usedFirstRow + Convert.ToInt32(usedRows.Count) - 1;
            var usedMaximumColumn = usedFirstColumn + Convert.ToInt32(usedColumns.Count) - 1;

            var readBounds = GetReadBounds(
                (object)worksheet,
                usedFirstRow,
                usedFirstColumn,
                usedMaximumRow,
                usedMaximumColumn);

            var model = new SheetModel {
                SheetName = Convert.ToString(worksheet.Name) ?? string.Empty,
                FirstRow = readBounds.FirstRow,
                FirstColumn = readBounds.FirstColumn,
                MaxRow = readBounds.MaximumRow,
                MaxColumn = readBounds.MaximumColumn
            };

            ReadColumnWidths((object)worksheet, model);
            ReadRowHeights((object)worksheet, model);
            ReadCells((object)worksheet, model);

            return model;
        }
        finally {
            ReleaseComObject(usedColumns);
            ReleaseComObject(usedRows);
            ReleaseComObject(usedRange);
            ReleaseComObject(worksheet);
        }
    }

    public SheetModel RefreshCellValues(ExcelSessionService session, SheetModel model) {
        dynamic? worksheet = null;
        dynamic? cells = null;

        try {
            worksheet = session.GetFirstWorksheet();
            cells = worksheet.Cells;

            var values = ReadRangeValues((object)worksheet, (object)cells, model);

            foreach (var cellModel in model.Cells.Where(cell => !cell.IsMergedSlave)) {
                var rawValue = GetRangeValue(
                    values,
                    cellModel.Row - model.FirstRow + 1,
                    cellModel.Column - model.FirstColumn + 1);

                cellModel.RawValue = rawValue;

                if (!HasBulkContent(rawValue)) {
                    cellModel.DisplayText = string.Empty;
                    continue;
                }

                dynamic? cell = null;

                try {
                    cell = cells[cellModel.Row, cellModel.Column];
                    ReadDisplayText((object)cell, cellModel);
                }
                finally {
                    ReleaseComObject(cell);
                }
            }

            return model;
        }
        finally {
            ReleaseComObject(cells);
            ReleaseComObject(worksheet);
        }
    }

    private static ReadBounds GetReadBounds(
        object worksheetObject,
        int usedFirstRow,
        int usedFirstColumn,
        int usedMaximumRow,
        int usedMaximumColumn) {
        var rowCount = usedMaximumRow - usedFirstRow + 1;
        var columnCount = usedMaximumColumn - usedFirstColumn + 1;
        var usedCellCount = (long)rowCount * columnCount;

        if (usedCellCount <= MaximumDirectReadCellCount)
            return new ReadBounds(
                usedFirstRow,
                usedFirstColumn,
                usedMaximumRow,
                usedMaximumColumn);

        var lastContentRow = FindLastContentCoordinate(
            worksheetObject,
            ExcelSearchOrderByRows,
            true);

        var lastContentColumn = FindLastContentCoordinate(
            worksheetObject,
            ExcelSearchOrderByColumns,
            false);

        if (lastContentRow == null || lastContentColumn == null)
            return new ReadBounds(
                usedFirstRow,
                usedFirstColumn,
                Math.Min(usedMaximumRow, usedFirstRow + 30),
                Math.Min(usedMaximumColumn, usedFirstColumn + 20));

        return new ReadBounds(
            usedFirstRow,
            usedFirstColumn,
            Math.Min(usedMaximumRow, lastContentRow.Value + AdditionalRowsAfterLastContent),
            Math.Min(usedMaximumColumn, lastContentColumn.Value + AdditionalColumnsAfterLastContent));
    }

    private static int? FindLastContentCoordinate(
        object worksheetObject,
        int searchOrder,
        bool returnRowCoordinate) {
        dynamic worksheet = worksheetObject;
        dynamic? cells = null;
        dynamic? foundCell = null;

        try {
            cells = worksheet.Cells;

            foundCell = cells.Find(
                "*",
                Type.Missing,
                ExcelFindLookInFormulas,
                Type.Missing,
                searchOrder,
                ExcelSearchDirectionPrevious,
                false,
                Type.Missing,
                Type.Missing);

            if (foundCell == null)
                return null;

            return returnRowCoordinate
                ? Convert.ToInt32(foundCell.Row)
                : Convert.ToInt32(foundCell.Column);
        }
        catch {
            return null;
        }
        finally {
            ReleaseComObject(foundCell);
            ReleaseComObject(cells);
        }
    }

    private static void ReadColumnWidths(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic? columns = null;

        try {
            columns = worksheet.Columns;

            for (var column = model.FirstColumn; column <= model.MaxColumn; column++) {
                dynamic? columnRange = null;

                try {
                    columnRange = columns[column];

                    var isHidden = Convert.ToBoolean(columnRange.Hidden);
                    var width = Convert.ToDouble(columnRange.Width);

                    model.ColumnWidths[column] = isHidden
                        ? 0.0
                        : Math.Max(width * PointsToDips, 0.0);
                }
                catch {
                    model.ColumnWidths[column] = model.DefaultColumnWidth;
                }
                finally {
                    ReleaseComObject(columnRange);
                }
            }
        }
        finally {
            ReleaseComObject(columns);
        }
    }

    private static void ReadRowHeights(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic? rows = null;

        try {
            rows = worksheet.Rows;

            for (var row = model.FirstRow; row <= model.MaxRow; row++) {
                dynamic? rowRange = null;

                try {
                    rowRange = rows[row];

                    var isHidden = Convert.ToBoolean(rowRange.Hidden);
                    var height = Convert.ToDouble(rowRange.Height);

                    model.RowHeights[row] = isHidden
                        ? 0.0
                        : Math.Max(height * PointsToDips, 0.0);
                }
                catch {
                    model.RowHeights[row] = model.DefaultRowHeight;
                }
                finally {
                    ReleaseComObject(rowRange);
                }
            }
        }
        finally {
            ReleaseComObject(rows);
        }
    }

    private void ReadCells(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        var mergedSlaveCells = new HashSet<(int Row, int Column)>();
        dynamic? cells = null;

        try {
            cells = worksheet.Cells;
            var values = ReadRangeValues(worksheetObject, (object)cells, model);

            for (var row = model.FirstRow; row <= model.MaxRow; row++) {
                if (model.RowHeights.TryGetValue(row, out var rowHeight) && rowHeight == 0.0)
                    continue;

                for (var column = model.FirstColumn; column <= model.MaxColumn; column++) {
                    if (model.ColumnWidths.TryGetValue(column, out var columnWidth) && columnWidth == 0.0)
                        continue;

                    if (mergedSlaveCells.Contains((row, column))) {
                        model.Cells.Add(new CellModel {
                            Row = row,
                            Column = column,
                            IsMergedSlave = true
                        });

                        continue;
                    }

                    var rawValue = GetRangeValue(
                        values,
                        row - model.FirstRow + 1,
                        column - model.FirstColumn + 1);

                    dynamic? cell = null;

                    try {
                        cell = cells[row, column];

                        var cellModel = ReadRenderableCell(
                            (object)cell,
                            row,
                            column,
                            worksheetObject,
                            rawValue);

                        if (cellModel == null)
                            continue;

                        RegisterMergedSlaveCells(cellModel, mergedSlaveCells);
                        model.Cells.Add(cellModel);
                    }
                    finally {
                        ReleaseComObject(cell);
                    }
                }
            }
        }
        finally {
            ReleaseComObject(cells);
        }
    }

    private CellModel? ReadRenderableCell(
        object cellObject,
        int row,
        int column,
        object worksheetObject,
        object? rawValue) {
        var model = new CellModel {
            Row = row,
            Column = column,
            RawValue = rawValue
        };

        if (HasBulkContent(rawValue))
            ReadDisplayText(cellObject, model);

        var hasBackground = ReadBackground(cellObject, model);

        if (!HasContent(model) && !hasBackground)
            return null;

        ReadFont(cellObject, model);
        ReadAlignment(cellObject, model);
        ReadBorders(cellObject, model);
        ReadMerge(cellObject, model);

        if (model.IsInput)
            ReadValidation(cellObject, model, worksheetObject);

        return model;
    }

    private static void RegisterMergedSlaveCells(
        CellModel cellModel,
        HashSet<(int Row, int Column)> mergedSlaveCells) {
        if (cellModel.RowSpan <= 1 && cellModel.ColSpan <= 1)
            return;

        for (var row = cellModel.Row; row < cellModel.Row + cellModel.RowSpan; row++)
        for (var column = cellModel.Column; column < cellModel.Column + cellModel.ColSpan; column++)
            if (row != cellModel.Row || column != cellModel.Column)
                mergedSlaveCells.Add((row, column));
    }

    private static bool HasContent(CellModel model) {
        return model.RawValue != null || !string.IsNullOrWhiteSpace(model.DisplayText);
    }

    private static object? ReadRangeValues(
        object worksheetObject,
        object cellsObject,
        SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic cells = cellsObject;
        dynamic? firstCell = null;
        dynamic? lastCell = null;
        dynamic? range = null;

        try {
            firstCell = cells[model.FirstRow, model.FirstColumn];
            lastCell = cells[model.MaxRow, model.MaxColumn];
            range = worksheet.Range[firstCell, lastCell];

            return range.Value2;
        }
        finally {
            ReleaseComObject(range);
            ReleaseComObject(lastCell);
            ReleaseComObject(firstCell);
        }
    }

    private static object? GetRangeValue(object? values, int rowOffset, int columnOffset) {
        if (values is not Array valuesArray || valuesArray.Rank != 2)
            return rowOffset == 1 && columnOffset == 1 ? values : null;

        var rowIndex = valuesArray.GetLowerBound(0) + rowOffset - 1;
        var columnIndex = valuesArray.GetLowerBound(1) + columnOffset - 1;

        if (rowIndex > valuesArray.GetUpperBound(0) || columnIndex > valuesArray.GetUpperBound(1))
            return null;

        return valuesArray.GetValue(rowIndex, columnIndex);
    }

    private static bool HasBulkContent(object? rawValue) {
        return rawValue switch {
            null => false,
            string text => !string.IsNullOrWhiteSpace(text),
            _ => true
        };
    }

    private static void ReadDisplayText(object cellObject, CellModel model) {
        dynamic cell = cellObject;

        try {
            model.DisplayText = Convert.ToString(cell.Text)
                                ?? Convert.ToString(model.RawValue)
                                ?? string.Empty;
        }
        catch {
            model.DisplayText = Convert.ToString(model.RawValue) ?? string.Empty;
        }
    }

    private static bool ReadBackground(object cellObject, CellModel model) {
        dynamic cell = cellObject;
        dynamic? interior = null;

        try {
            interior = cell.Interior;
            var colorIndex = Convert.ToInt32(interior.ColorIndex);

            if (colorIndex == ExcelColorIndexNone) {
                model.BackgroundColor = Colors.White;
                model.IsInput = false;
                model.InputType = CellInputType.None;
                return false;
            }

            model.BackgroundColor = OleColorToMediaColor(Convert.ToInt32(interior.Color));

            var colorHex = ToHex(model.BackgroundColor);

            model.IsInput = colorHex is InputColorLightGreen or InputColorGreen or DropdownColor;
            model.InputType = colorHex == DropdownColor
                ? CellInputType.ComboBox
                : model.IsInput
                    ? CellInputType.TextBox
                    : CellInputType.None;

            return true;
        }
        catch {
            model.BackgroundColor = Colors.White;
            model.IsInput = false;
            model.InputType = CellInputType.None;
            return false;
        }
        finally {
            ReleaseComObject(interior);
        }
    }

    private static void ReadFont(object cellObject, CellModel model) {
        dynamic cell = cellObject;
        dynamic? font = null;

        try {
            font = cell.Font;
            model.IsBold = Convert.ToBoolean(font.Bold);
            model.IsItalic = Convert.ToBoolean(font.Italic);

            try {
                model.FontSize = Convert.ToDouble(font.Size);
            }
            catch {
                model.FontSize = 11.0;
            }

            try {
                model.ForegroundColor = font.Color == null
                    ? Colors.Black
                    : OleColorToMediaColor(Convert.ToInt32(font.Color));
            }
            catch {
                model.ForegroundColor = Colors.Black;
            }
        }
        catch {
        }
        finally {
            ReleaseComObject(font);
        }
    }

    private static void ReadAlignment(object cellObject, CellModel model) {
        dynamic cell = cellObject;

        try {
            model.TextAlignment = ExcelHorizontalAlignmentToWpf(cell.HorizontalAlignment);
            model.VerticalContentAlignment = ExcelVerticalAlignmentToWpf(cell.VerticalAlignment);
            model.WrapText = Convert.ToBoolean(cell.WrapText);
        }
        catch {
        }
    }

    private void ReadBorders(object cellObject, CellModel model) {
        dynamic cell = cellObject;
        dynamic? borders = null;

        try {
            borders = cell.Borders;
            var dominantColor = Colors.Black;

            model.BorderTopThickness = ReadBorderThickness(borders, ExcelBorderEdgeTop, ref dominantColor);
            model.BorderBottomThickness = ReadBorderThickness(borders, ExcelBorderEdgeBottom, ref dominantColor);
            model.BorderLeftThickness = ReadBorderThickness(borders, ExcelBorderEdgeLeft, ref dominantColor);
            model.BorderRightThickness = ReadBorderThickness(borders, ExcelBorderEdgeRight, ref dominantColor);
            model.BorderColor = dominantColor;
        }
        catch {
        }
        finally {
            ReleaseComObject(borders);
        }
    }

    private static double ReadBorderThickness(
        dynamic borders,
        int borderIndex,
        ref Color dominantColor) {
        dynamic? border = null;

        try {
            border = borders[borderIndex];

            if (Convert.ToInt32(border.LineStyle) == ExcelLineStyleNone)
                return 0.0;

            try {
                dominantColor = OleColorToMediaColor(Convert.ToInt32(border.Color));
            }
            catch {
            }

            return ExcelBorderWeightToDips(border.Weight);
        }
        catch {
            return 0.0;
        }
        finally {
            ReleaseComObject(border);
        }
    }

    private static void ReadMerge(object cellObject, CellModel model) {
        dynamic cell = cellObject;
        dynamic? mergeArea = null;
        dynamic? rows = null;
        dynamic? columns = null;

        try {
            if (!Convert.ToBoolean(cell.MergeCells))
                return;

            mergeArea = cell.MergeArea;
            rows = mergeArea.Rows;
            columns = mergeArea.Columns;

            model.RowSpan = Convert.ToInt32(rows.Count);
            model.ColSpan = Convert.ToInt32(columns.Count);
        }
        catch {
        }
        finally {
            ReleaseComObject(columns);
            ReleaseComObject(rows);
            ReleaseComObject(mergeArea);
        }
    }

    private void ReadValidation(
        object cellObject,
        CellModel model,
        object worksheetObject) {
        dynamic cell = cellObject;
        dynamic? validation = null;

        try {
            validation = cell.Validation;

            if (Convert.ToInt32(validation.Type) != ExcelValidationTypeList)
                return;

            var formula = Convert.ToString(validation.Formula1);

            if (string.IsNullOrWhiteSpace(formula))
                return;

            var values = ParseValidationFormula(formula, worksheetObject);

            if (values.Count == 0)
                return;

            model.DropdownValues = values;
            model.InputType = CellInputType.ComboBox;
        }
        catch {
        }
        finally {
            ReleaseComObject(validation);
        }
    }

    private static List<string> ParseValidationFormula(
        string formula,
        object worksheetObject) {
        dynamic worksheet = worksheetObject;
        var values = new List<string>();

        if (!formula.StartsWith('=')) {
            values.AddRange(
                formula
                    .Split(',')
                    .Select(value => value.Trim())
                    .Where(value => !string.IsNullOrWhiteSpace(value)));

            return values;
        }

        dynamic? range = null;
        dynamic? cells = null;

        try {
            range = worksheet.Range[formula.TrimStart('=')];
            cells = range.Cells;

            var count = Convert.ToInt32(cells.Count);

            for (var index = 1; index <= count; index++) {
                dynamic? rangeCell = null;

                try {
                    rangeCell = cells[index];
                    var value = Convert.ToString(rangeCell.Value2);

                    if (!string.IsNullOrWhiteSpace(value))
                        values.Add(value);
                }
                finally {
                    ReleaseComObject(rangeCell);
                }
            }
        }
        catch {
        }
        finally {
            ReleaseComObject(cells);
            ReleaseComObject(range);
        }

        return values;
    }

    private static string ToHex(Color color) {
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    }

    private static Color OleColorToMediaColor(int oleColor) {
        var red = (byte)(oleColor & 0xFF);
        var green = (byte)((oleColor >> 8) & 0xFF);
        var blue = (byte)((oleColor >> 16) & 0xFF);

        return Color.FromRgb(red, green, blue);
    }

    private static TextAlignment ExcelHorizontalAlignmentToWpf(object horizontalAlignment) {
        int value;

        try {
            value = Convert.ToInt32(horizontalAlignment);
        }
        catch {
            return TextAlignment.Left;
        }

        return value switch {
            ExcelHorizontalAlignmentLeft => TextAlignment.Left,
            ExcelHorizontalAlignmentCenter => TextAlignment.Center,
            ExcelHorizontalAlignmentRight => TextAlignment.Right,
            ExcelHorizontalAlignmentJustify => TextAlignment.Justify,
            ExcelHorizontalAlignmentGeneral => TextAlignment.Left,
            _ => TextAlignment.Left
        };
    }

    private static VerticalAlignment ExcelVerticalAlignmentToWpf(object verticalAlignment) {
        int value;

        try {
            value = Convert.ToInt32(verticalAlignment);
        }
        catch {
            return VerticalAlignment.Center;
        }

        return value switch {
            ExcelVerticalAlignmentTop => VerticalAlignment.Top,
            ExcelVerticalAlignmentCenter => VerticalAlignment.Center,
            ExcelVerticalAlignmentBottom => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Center
        };
    }

    private static double ExcelBorderWeightToDips(object borderWeight) {
        int value;

        try {
            value = Convert.ToInt32(borderWeight);
        }
        catch {
            return 1.0;
        }

        return value switch {
            ExcelBorderWeightHairline => 0.5,
            ExcelBorderWeightThin => 1.0,
            ExcelBorderWeightMedium => 2.0,
            ExcelBorderWeightThick => 3.0,
            _ => 1.0
        };
    }

    private static void ReleaseComObject(object? comObject) {
        if (comObject == null || !Marshal.IsComObject(comObject))
            return;

        try {
            Marshal.FinalReleaseComObject(comObject);
        }
        catch {
        }
    }

    private sealed record ReadBounds(
        int FirstRow,
        int FirstColumn,
        int MaximumRow,
        int MaximumColumn);
}