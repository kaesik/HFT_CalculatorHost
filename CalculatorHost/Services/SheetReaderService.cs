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
    private const int ExcelCellTypeAllValidation = -4174;
    private const int ExcelReferenceStyleA1 = 1;
    private const int ExcelReferenceTypeAbsolute = 1;
    private const int ExcelShapeTypeFormControl = 8;
    private const int ExcelFormControlTypeDropdown = 2;
    private const int ExcelFormControlTypeListBox = 6;
    private const int ExcelCellControlTypeNone = 0;
    private const int ExcelCellControlTypeCheckbox = 2;

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
            ReadDropdownElements((object)worksheet, model);

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

            ReadDropdownElements((object)worksheet, model);

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

    private void ReadDropdownElements(object worksheetObject, SheetModel model) {
        ReadDataValidationDropdowns(worksheetObject, model);
        ReadFormControlDropdowns(worksheetObject, model);
        ReadActiveXDropdowns(worksheetObject, model);
        ReadCellControlDropdowns(worksheetObject, model);
    }

    private void ReadCellControlDropdowns(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic? cells = null;

        try {
            cells = worksheet.Cells;

            foreach (var cellModel in model.Cells.Where(cell =>
                         !cell.IsMergedSlave &&
                         cell.InputType == CellInputType.ComboBox &&
                         cell.DropdownValues.Count == 0)) {
                dynamic? cell = null;
                dynamic? cellControl = null;

                try {
                    cell = cells[cellModel.Row, cellModel.Column];
                    cellControl = cell.CellControl;

                    var type = Convert.ToInt32(cellControl.Type);

                    if (type == ExcelCellControlTypeNone || type == ExcelCellControlTypeCheckbox)
                        continue;

                    var values = ReadCellControlValues(
                        (object)cellControl,
                        (object)cell,
                        worksheetObject);

                    if (values.Count == 0)
                        continue;

                    ApplyDropdownValues(
                        cellModel,
                        values,
                        Convert.ToString(cell.Value2),
                        null,
                        false);
                }
                catch {
                }
                finally {
                    ReleaseComObject(cellControl);
                    ReleaseComObject(cell);
                }
            }
        }
        catch {
        }
        finally {
            ReleaseComObject(cells);
        }
    }

    private static List<string> ReadCellControlValues(
        object cellControlObject,
        object cellObject,
        object worksheetObject) {
        dynamic cellControl = cellControlObject;
        dynamic cell = cellObject;

        try {
            var fillRange = Convert.ToString(cellControl.ListFillRange);
            var values = ReadDropdownSourceValues(fillRange, worksheetObject);

            if (values.Count > 0)
                return values;
        }
        catch {
        }

        try {
            var source = Convert.ToString(cellControl.Source);
            var values = ReadDropdownSourceValues(source, worksheetObject);

            if (values.Count > 0)
                return values;
        }
        catch {
        }

        try {
            var values = ExtractValidationValues(cellControl.Items);

            if (values.Count > 0)
                return values;
        }
        catch {
        }

        try {
            var values = ExtractValidationValues(cellControl.Values);

            if (values.Count > 0)
                return values;
        }
        catch {
        }

        try {
            var values = ExtractValidationValues(cell.Validation.Formula1);

            if (values.Count > 0)
                return values;
        }
        catch {
        }

        return [];
    }

    private void ReadDataValidationDropdowns(object worksheetObject, SheetModel model) {
        ReadRenderedInputValidations(worksheetObject, model);

        dynamic worksheet = worksheetObject;
        dynamic? validationCells = null;
        dynamic? cells = null;

        try {
            validationCells = worksheet.Cells.SpecialCells(ExcelCellTypeAllValidation);
            cells = validationCells.Cells;

            var count = Convert.ToInt32(cells.Count);

            for (var index = 1; index <= count; index++) {
                dynamic? cell = null;

                try {
                    cell = cells[index];

                    var row = Convert.ToInt32(cell.Row);
                    var column = Convert.ToInt32(cell.Column);

                    if (!IsInsideRenderedSheet(row, column, model))
                        continue;

                    var cellModel = GetOrCreateDropdownCellModel(
                        (object)cell,
                        row,
                        column,
                        model);

                    ReadValidation((object)cell, cellModel, worksheetObject);
                }
                finally {
                    ReleaseComObject(cell);
                }
            }
        }
        catch {
        }
        finally {
            ReleaseComObject(cells);
            ReleaseComObject(validationCells);
        }
    }

    private void ReadRenderedInputValidations(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic? cells = null;

        try {
            cells = worksheet.Cells;

            foreach (var cellModel in model.Cells.Where(cell =>
                         !cell.IsMergedSlave &&
                         cell.IsInput)) {
                dynamic? cell = null;

                try {
                    cell = cells[cellModel.Row, cellModel.Column];
                    ReadValidation((object)cell, cellModel, worksheetObject);
                }
                catch {
                }
                finally {
                    ReleaseComObject(cell);
                }
            }
        }
        catch {
        }
        finally {
            ReleaseComObject(cells);
        }
    }

    private void ReadFormControlDropdowns(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic? shapes = null;

        try {
            shapes = worksheet.Shapes;
            var count = Convert.ToInt32(shapes.Count);

            for (var index = 1; index <= count; index++) {
                dynamic? shape = null;
                dynamic? controlFormat = null;
                dynamic? topLeftCell = null;

                try {
                    shape = shapes.Item(index);

                    if (Convert.ToInt32(shape.Type) != ExcelShapeTypeFormControl)
                        continue;

                    var controlType = Convert.ToInt32(shape.FormControlType);

                    if (controlType != ExcelFormControlTypeDropdown &&
                        controlType != ExcelFormControlTypeListBox)
                        continue;

                    controlFormat = shape.ControlFormat;
                    var values = ReadFormControlValues((object)controlFormat, worksheetObject);

                    if (values.Count == 0)
                        continue;

                    topLeftCell = shape.TopLeftCell;

                    var row = Convert.ToInt32(topLeftCell.Row);
                    var column = Convert.ToInt32(topLeftCell.Column);

                    if (!IsInsideRenderedSheet(row, column, model))
                        continue;

                    var cellModel = GetOrCreateDropdownCellModel(
                        (object)topLeftCell,
                        row,
                        column,
                        model);

                    var selectedValue = ReadSelectedFormControlValue((object)controlFormat, values);
                    var inputTarget = ReadLinkedCellPosition((object)controlFormat, worksheetObject);

                    ApplyDropdownValues(
                        cellModel,
                        values,
                        selectedValue,
                        inputTarget,
                        inputTarget != null);
                }
                catch {
                }
                finally {
                    ReleaseComObject(topLeftCell);
                    ReleaseComObject(controlFormat);
                    ReleaseComObject(shape);
                }
            }
        }
        catch {
        }
        finally {
            ReleaseComObject(shapes);
        }
    }

    private void ReadActiveXDropdowns(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic? objects = null;

        try {
            objects = worksheet.OLEObjects();
            var count = Convert.ToInt32(objects.Count);

            for (var index = 1; index <= count; index++) {
                dynamic? embeddedObject = null;
                dynamic? control = null;
                dynamic? topLeftCell = null;

                try {
                    embeddedObject = objects.Item(index);
                    control = embeddedObject.Object;

                    var values = ReadActiveXValues(
                        (object)embeddedObject,
                        (object)control,
                        worksheetObject);

                    if (values.Count == 0)
                        continue;

                    topLeftCell = embeddedObject.TopLeftCell;

                    var row = Convert.ToInt32(topLeftCell.Row);
                    var column = Convert.ToInt32(topLeftCell.Column);

                    if (!IsInsideRenderedSheet(row, column, model))
                        continue;

                    var cellModel = GetOrCreateDropdownCellModel(
                        (object)topLeftCell,
                        row,
                        column,
                        model);

                    var selectedValue = ReadActiveXSelectedValue((object)control);
                    var inputTarget = ReadLinkedCellPosition((object)embeddedObject, worksheetObject)
                                      ?? ReadLinkedCellPosition((object)control, worksheetObject);

                    ApplyDropdownValues(
                        cellModel,
                        values,
                        selectedValue,
                        inputTarget,
                        false);
                }
                catch {
                }
                finally {
                    ReleaseComObject(topLeftCell);
                    ReleaseComObject(control);
                    ReleaseComObject(embeddedObject);
                }
            }
        }
        catch {
        }
        finally {
            ReleaseComObject(objects);
        }
    }

    private CellModel GetOrCreateDropdownCellModel(
        object cellObject,
        int row,
        int column,
        SheetModel model) {
        var existingCell = model.Cells.FirstOrDefault(cell =>
            cell.Row == row &&
            cell.Column == column &&
            !cell.IsMergedSlave);

        if (existingCell != null)
            return existingCell;

        dynamic cell = cellObject;

        var cellModel = new CellModel {
            Row = row,
            Column = column
        };

        try {
            cellModel.RawValue = cell.Value2;
        }
        catch {
        }

        ReadDisplayText(cellObject, cellModel);
        ReadBackground(cellObject, cellModel);
        ReadFont(cellObject, cellModel);
        ReadAlignment(cellObject, cellModel);
        ReadBorders(cellObject, cellModel);
        ReadMerge(cellObject, cellModel);

        model.Cells.Add(cellModel);

        return cellModel;
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

            var formulas = new List<string>();

            TryAddValidationFormula(formulas, Convert.ToString(validation.Formula1));

            try {
                TryAddValidationFormula(formulas, Convert.ToString(validation.Formula1Local));
            }
            catch {
            }

            foreach (var formula in formulas) {
                var values = ReadDropdownSourceValues(formula, worksheetObject, cellObject);

                if (values.Count == 0)
                    continue;

                ApplyDropdownValues(model, values, null, null, false);
                return;
            }
        }
        catch {
        }
        finally {
            ReleaseComObject(validation);
        }
    }

    private static void TryAddValidationFormula(List<string> formulas, string? formula) {
        var normalizedFormula = formula?.Trim();

        if (string.IsNullOrWhiteSpace(normalizedFormula))
            return;

        if (!formulas.Contains(normalizedFormula, StringComparer.Ordinal))
            formulas.Add(normalizedFormula);

        var dynamicArrayFormula = NormalizeDynamicArrayFormula(normalizedFormula);

        if (!formulas.Contains(dynamicArrayFormula, StringComparer.Ordinal))
            formulas.Add(dynamicArrayFormula);
    }

    private static string NormalizeDynamicArrayFormula(string formula) {
        const string anchorArrayFunction = "_xlfn.ANCHORARRAY(";

        var functionStart = formula.IndexOf(anchorArrayFunction, StringComparison.OrdinalIgnoreCase);

        if (functionStart < 0 || !formula.EndsWith(')'))
            return formula;

        var argumentStart = functionStart + anchorArrayFunction.Length;
        var argumentLength = formula.Length - argumentStart - 1;

        if (argumentLength <= 0)
            return formula;

        var argument = formula.Substring(argumentStart, argumentLength);

        return $"{formula[..functionStart]}{argument}#";
    }

    private static List<string> ReadDropdownSourceValues(
        string? source,
        object worksheetObject,
        object? relativeToCellObject = null) {
        if (string.IsNullOrWhiteSpace(source))
            return [];

        var trimmedSource = source.Trim();

        if (!trimmedSource.StartsWith('=') &&
            (trimmedSource.Contains(',') || trimmedSource.Contains(';')))
            return SplitValidationValues(trimmedSource);

        foreach (var expression in BuildEvaluationExpressions(trimmedSource)) {
            var expressionInCellContext = ConvertRelativeValidationFormula(
                expression,
                worksheetObject,
                relativeToCellObject);

            var values = ReadEvaluatedValidationValues(expressionInCellContext, worksheetObject);

            if (values.Count > 0)
                return values;

            if (!string.Equals(expressionInCellContext, expression, StringComparison.Ordinal)) {
                values = ReadEvaluatedValidationValues(expression, worksheetObject);

                if (values.Count > 0)
                    return values;
            }
        }

        var reference = trimmedSource.TrimStart('=');
        var rangeValues = ReadRangeValidationValues(reference, worksheetObject);

        if (rangeValues.Count > 0)
            return rangeValues;

        return trimmedSource.StartsWith('=')
            ? []
            : SplitValidationValues(trimmedSource);
    }

    private static string ConvertRelativeValidationFormula(
        string formula,
        object worksheetObject,
        object? relativeToCellObject) {
        if (relativeToCellObject == null || !formula.StartsWith('='))
            return formula;

        dynamic worksheet = worksheetObject;
        dynamic? application = null;

        try {
            application = worksheet.Application;

            var convertedFormula = application.ConvertFormula(
                formula,
                ExcelReferenceStyleA1,
                ExcelReferenceStyleA1,
                ExcelReferenceTypeAbsolute,
                relativeToCellObject);

            return Convert.ToString(convertedFormula) ?? formula;
        }
        catch {
            return formula;
        }
        finally {
            ReleaseComObject(application);
        }
    }

    private static List<string> BuildEvaluationExpressions(string source) {
        var expressions = new List<string>();

        TryAddExpression(expressions, source.StartsWith('=') ? source : $"={source}");
        TryAddExpression(expressions, NormalizeDynamicArrayFormula(
            source.StartsWith('=') ? source : $"={source}"));

        return expressions;
    }

    private static void TryAddExpression(List<string> expressions, string expression) {
        if (!expressions.Contains(expression, StringComparer.Ordinal))
            expressions.Add(expression);
    }

    private static List<string> SplitValidationValues(string valuesText) {
        var normalizedText = valuesText.Trim();

        if (normalizedText.Length >= 2 &&
            normalizedText.StartsWith('"') &&
            normalizedText.EndsWith('"'))
            normalizedText = normalizedText[1..^1];

        return normalizedText
            .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(value => value.Trim())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.CurrentCulture)
            .ToList();
    }

    private static List<string> ReadEvaluatedValidationValues(
        string formula,
        object worksheetObject) {
        dynamic worksheet = worksheetObject;
        dynamic? application = null;
        object? evaluatedValue = null;

        try {
            try {
                evaluatedValue = worksheet.Evaluate(formula);
                var values = ExtractValidationValues(evaluatedValue);

                if (values.Count > 0)
                    return values;
            }
            catch {
            }
            finally {
                ReleaseComObject(evaluatedValue);
                evaluatedValue = null;
            }

            application = worksheet.Application;
            evaluatedValue = application.Evaluate(formula);

            return ExtractValidationValues(evaluatedValue);
        }
        catch {
            return [];
        }
        finally {
            ReleaseComObject(evaluatedValue);
            ReleaseComObject(application);
        }
    }

    private static List<string> ReadRangeValidationValues(
        string reference,
        object worksheetObject) {
        dynamic worksheet = worksheetObject;
        dynamic? range = null;

        try {
            range = worksheet.Range[reference];
            var values = ExtractValidationValues(range);

            if (values.Count > 0)
                return values;
        }
        catch {
        }
        finally {
            ReleaseComObject(range);
            range = null;
        }

        foreach (var expression in BuildEvaluationExpressions(reference)) {
            var values = ReadEvaluatedValidationValues(expression, worksheetObject);

            if (values.Count > 0)
                return values;
        }

        return [];
    }

    private static List<string> ReadFormControlValues(
        object controlFormatObject,
        object worksheetObject) {
        dynamic controlFormat = controlFormatObject;
        var values = ReadControlFillRangeValues(controlFormatObject, worksheetObject);

        if (values.Count > 0)
            return values;

        try {
            var count = Convert.ToInt32(controlFormat.ListCount);

            for (var index = 1; index <= count; index++)
                AddDropdownValue(values, Convert.ToString(controlFormat.List[index]));
        }
        catch {
        }

        return values;
    }

    private static List<string> ReadActiveXValues(
        object embeddedObjectObject,
        object controlObject,
        object worksheetObject) {
        var values = ReadControlFillRangeValues(embeddedObjectObject, worksheetObject);

        if (values.Count > 0)
            return values;

        values = ReadControlFillRangeValues(controlObject, worksheetObject);

        if (values.Count > 0)
            return values;

        dynamic control = controlObject;

        try {
            values = ExtractValidationValues(control.List);

            if (values.Count > 0)
                return values;
        }
        catch {
        }

        try {
            var count = Convert.ToInt32(control.ListCount);

            for (var index = 0; index < count; index++)
                AddDropdownValue(values, Convert.ToString(control.List[index]));
        }
        catch {
        }

        return values;
    }

    private static List<string> ReadControlFillRangeValues(
        object controlObject,
        object worksheetObject) {
        dynamic control = controlObject;

        try {
            var fillRange = Convert.ToString(control.ListFillRange);
            return ReadDropdownSourceValues(fillRange, worksheetObject);
        }
        catch {
            return [];
        }
    }

    private static string? ReadSelectedFormControlValue(
        object controlFormatObject,
        IReadOnlyList<string> values) {
        dynamic controlFormat = controlFormatObject;

        try {
            var selectedIndex = Convert.ToInt32(controlFormat.Value);

            return selectedIndex > 0 && selectedIndex <= values.Count
                ? values[selectedIndex - 1]
                : null;
        }
        catch {
            return null;
        }
    }

    private static string? ReadActiveXSelectedValue(object controlObject) {
        dynamic control = controlObject;

        try {
            return Convert.ToString(control.Value);
        }
        catch {
            return null;
        }
    }

    private static (int Row, int Column)? ReadLinkedCellPosition(
        object controlObject,
        object worksheetObject) {
        dynamic control = controlObject;
        dynamic worksheet = worksheetObject;
        dynamic? linkedCell = null;

        try {
            var linkedCellReference = Convert.ToString(control.LinkedCell);

            if (string.IsNullOrWhiteSpace(linkedCellReference))
                return null;

            linkedCell = worksheet.Range[linkedCellReference];

            return (
                Convert.ToInt32(linkedCell.Row),
                Convert.ToInt32(linkedCell.Column));
        }
        catch {
            return null;
        }
        finally {
            ReleaseComObject(linkedCell);
        }
    }

    private static void ApplyDropdownValues(
        CellModel model,
        IEnumerable<string> values,
        string? selectedValue,
        (int Row, int Column)? inputTarget,
        bool dropdownWritesSelectedIndex) {
        model.IsInput = true;
        model.InputType = CellInputType.ComboBox;
        model.DropdownValues = values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.CurrentCulture)
            .ToList();

        model.InputTargetRow = inputTarget?.Row;
        model.InputTargetColumn = inputTarget?.Column;
        model.DropdownWritesSelectedIndex = dropdownWritesSelectedIndex;

        if (!string.IsNullOrWhiteSpace(selectedValue))
            model.DisplayText = selectedValue;
    }

    private static List<string> ExtractValidationValues(object? source) {
        var values = new List<string>();

        if (source == null)
            return values;

        if (Marshal.IsComObject(source)) {
            dynamic range = source;

            try {
                return ExtractValidationValues(range.Value2);
            }
            catch {
                return values;
            }
        }

        if (source is Array array) {
            foreach (var item in array)
                AddDropdownValue(values, Convert.ToString(item));

            return values;
        }

        AddDropdownValue(values, Convert.ToString(source));

        return values;
    }

    private static void AddDropdownValue(List<string> values, string? value) {
        var normalizedValue = value?.Trim();

        if (string.IsNullOrWhiteSpace(normalizedValue))
            return;

        if (!values.Contains(normalizedValue, StringComparer.CurrentCulture))
            values.Add(normalizedValue);
    }

    private static bool IsInsideRenderedSheet(int row, int column, SheetModel model) {
        return row >= model.FirstRow &&
               row <= model.MaxRow &&
               column >= model.FirstColumn &&
               column <= model.MaxColumn;
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