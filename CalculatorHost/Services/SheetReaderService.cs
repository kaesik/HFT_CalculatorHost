using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public enum SheetRefreshMode {
    ValuesOnly,
    ValuesAndImages,
    Full
}

public sealed record SheetReadProgress(
    string Stage,
    double OverallPercentage,
    double StagePercentage,
    string Detail);

public class SheetReaderService {
    private const double PointsToDips = 96.0 / 72.0;

    private const int MaximumDirectReadCellCount = 250000;
    private const int AdditionalRowsAfterLastContent = 80;
    private const int AdditionalColumnsAfterLastContent = 40;

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
    private const int ExcelShapeTypeLine = 9;
    private const int ExcelShapeTypeLinkedPicture = 11;
    private const int ExcelShapeTypePicture = 13;
    private const int ExcelShapeTypeGraphic = 28;
    private const int ExcelShapeTypeLinkedGraphic = 29;
    private const int ExcelCopyPictureAppearanceScreen = 1;
    private const int ExcelCopyPictureFormatBitmap = 2;
    private const int MaximumClipboardReadAttempts = 5;
    private const int ClipboardRetryDelayMilliseconds = 20;
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

    private const string DefaultInputColor = "#92D050";
    private const string DefaultSecondInputColor = "#00B050";
    private const string DefaultDropdownColor = "#00B0F0";
    private const string DefaultOutputColor = "#FFFFFF";

    private const int RoleColorColumn = 13;
    private const int InputRoleColorRow = 1;
    private const int DropdownRoleColorRow = 2;
    private const int SecondInputRoleColorRow = 3;
    private const int OutputRoleColorRow = 4;

    public SheetModel ReadFirstSheet(
        ExcelSessionService session,
        IProgress<SheetReadProgress>? progress = null) {
        dynamic? worksheet = null;
        dynamic? usedRange = null;
        dynamic? usedRows = null;
        dynamic? usedColumns = null;

        try {
            ReportProgress(progress, "Analiza arkusza", 0.0, 0.0, "Ustalanie zakresu danych…");

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

            var rowCount = Math.Max(model.MaxRow - model.FirstRow + 1, 0);
            var columnCount = Math.Max(model.MaxColumn - model.FirstColumn + 1, 0);

            ReportProgress(
                progress,
                "Analiza arkusza",
                3.0,
                100.0,
                $"Zakres: {rowCount} w. × {columnCount} kol.");

            ReportProgress(progress, "Rozpoznawanie pól", 3.0, 0.0, "Odczyt kolorów pól…");
            var colorRoles = ReadColorRoles((object)worksheet);
            ReportProgress(progress, "Rozpoznawanie pól", 5.0, 100.0, "Gotowe");

            ReadColumnWidths((object)worksheet, model, progress);
            ReadRowHeights((object)worksheet, model, progress);
            ReadCells((object)worksheet, model, colorRoles, progress);
            ReadDropdownElements((object)worksheet, model, colorRoles, progress);
            ReadImages((object)worksheet, model, progress);
            ReadMacroButtons((object)worksheet, model, progress);

            ReportProgress(progress, "Gotowe", 100.0, 100.0, "Arkusz został wczytany.");

            return model;
        }
        finally {
            ReleaseComObject(usedColumns);
            ReleaseComObject(usedRows);
            ReleaseComObject(usedRange);
            ReleaseComObject(worksheet);
        }
    }

    public SheetModel RefreshCellValues(
        ExcelSessionService session,
        SheetModel model,
        SheetRefreshMode refreshMode = SheetRefreshMode.Full,
        IProgress<SheetReadProgress>? progress = null) {
        dynamic? worksheet = null;
        dynamic? cells = null;

        try {
            worksheet = session.GetFirstWorksheet();
            cells = worksheet.Cells;

            var values = ReadRangeValues((object)worksheet, (object)cells, model);
            var refreshableCells = model.Cells
                .Where(cell => !cell.IsMergedSlave)
                .ToList();

            var valuesOverallEnd = refreshMode switch {
                SheetRefreshMode.ValuesOnly => 100.0,
                SheetRefreshMode.ValuesAndImages => 80.0,
                _ => 75.0
            };

            ReportStageProgress(
                progress,
                "Odświeżanie wartości",
                0.0,
                valuesOverallEnd,
                0,
                refreshableCells.Count,
                refreshableCells.Count == 0
                    ? "Brak komórek do odświeżenia"
                    : $"Komórka 0 z {refreshableCells.Count}");

            var processedCells = 0;

            foreach (var cellModel in refreshableCells) {
                var rawValue = GetRangeValue(
                    values,
                    cellModel.Row - model.FirstRow + 1,
                    cellModel.Column - model.FirstColumn + 1);

                var valueChanged = !AreEquivalentCellValues(cellModel.RawValue, rawValue);

                cellModel.RawValue = rawValue;

                if (!HasBulkContent(rawValue))
                    cellModel.DisplayText = string.Empty;
                else if (valueChanged || string.IsNullOrEmpty(cellModel.DisplayText)) {
                    dynamic? cell = null;

                    try {
                        cell = cells[cellModel.Row, cellModel.Column];
                        ReadDisplayText((object)cell, cellModel);
                    }
                    finally {
                        ReleaseComObject(cell);
                    }
                }

                processedCells++;

                ReportStageProgress(
                    progress,
                    "Odświeżanie wartości",
                    0.0,
                    valuesOverallEnd,
                    processedCells,
                    refreshableCells.Count,
                    $"Komórka {processedCells} z {refreshableCells.Count}");
            }

            if (refreshMode == SheetRefreshMode.Full) {
                ReportProgress(progress, "Rozpoznawanie pól", 75.0, 0.0, "Odczyt kolorów pól…");
                var colorRoles = ReadColorRoles((object)worksheet);
                ReportProgress(progress, "Rozpoznawanie pól", 78.0, 100.0, "Gotowe");

                ReadDropdownElements((object)worksheet, model, colorRoles, progress, 78.0, 88.0);
                ReadImages((object)worksheet, model, progress, 88.0, 95.0);
                ReadMacroButtons((object)worksheet, model, progress, 95.0);
            }
            else if (refreshMode == SheetRefreshMode.ValuesAndImages)
                ReadImages((object)worksheet, model, progress, 80.0, 100.0);

            ReportProgress(progress, "Gotowe", 100.0, 100.0, "Wartości arkusza zostały odświeżone.");

            return model;
        }
        finally {
            ReleaseComObject(cells);
            ReleaseComObject(worksheet);
        }
    }

    private static bool AreEquivalentCellValues(object? previousValue, object? currentValue) {
        if (ReferenceEquals(previousValue, currentValue))
            return true;

        if (previousValue == null || currentValue == null)
            return false;

        if (previousValue.Equals(currentValue))
            return true;

        // Excel COM can return the same numeric value using different CLR numeric types.
        if (previousValue is IConvertible && currentValue is IConvertible)
            try {
                var previousNumber = Convert.ToDouble(previousValue, CultureInfo.InvariantCulture);
                var currentNumber = Convert.ToDouble(currentValue, CultureInfo.InvariantCulture);
                return previousNumber.Equals(currentNumber);
            }
            catch {
                // At least one value is not numeric; compare its invariant text below.
            }

        return string.Equals(
            Convert.ToString(previousValue, CultureInfo.InvariantCulture),
            Convert.ToString(currentValue, CultureInfo.InvariantCulture),
            StringComparison.Ordinal);
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
                Math.Min(usedMaximumRow, usedFirstRow + 160),
                Math.Min(usedMaximumColumn, usedFirstColumn + 80));

        return new ReadBounds(
            usedFirstRow,
            usedFirstColumn,
            Math.Min(usedMaximumRow,
                Math.Max(lastContentRow.Value + AdditionalRowsAfterLastContent, usedFirstRow + 160)),
            Math.Min(usedMaximumColumn,
                Math.Max(lastContentColumn.Value + AdditionalColumnsAfterLastContent, usedFirstColumn + 80)));
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

    private static CellColorRoles ReadColorRoles(object worksheetObject) {
        var roles = new CellColorRoles();

        roles.InputColors.Add(ReadCellBackgroundHex(worksheetObject, InputRoleColorRow, RoleColorColumn) ??
                              DefaultInputColor);
        roles.InputColors.Add(ReadCellBackgroundHex(worksheetObject, SecondInputRoleColorRow, RoleColorColumn) ??
                              DefaultSecondInputColor);
        roles.DropdownColors.Add(ReadCellBackgroundHex(worksheetObject, DropdownRoleColorRow, RoleColorColumn) ??
                                 DefaultDropdownColor);
        roles.OutputColors.Add(ReadCellBackgroundHex(worksheetObject, OutputRoleColorRow, RoleColorColumn) ??
                               DefaultOutputColor);

        return roles;
    }

    private static string? ReadCellBackgroundHex(object worksheetObject, int row, int column) {
        dynamic worksheet = worksheetObject;
        dynamic? cells = null;
        dynamic? cell = null;

        try {
            cells = worksheet.Cells;
            cell = cells[row, column];

            return ReadCellBackgroundHex((object)cell);
        }
        catch {
            return null;
        }
        finally {
            ReleaseComObject(cell);
            ReleaseComObject(cells);
        }
    }

    private static string? ReadCellBackgroundHex(object cellObject) {
        dynamic cell = cellObject;
        dynamic? interior = null;

        try {
            interior = cell.Interior;

            if (Convert.ToInt32(interior.ColorIndex) == ExcelColorIndexNone)
                return null;

            return ToHex(OleColorToMediaColor(Convert.ToInt32(interior.Color)));
        }
        catch {
            return null;
        }
        finally {
            ReleaseComObject(interior);
        }
    }

    private static bool IsRoleColorDefinitionCell(int row, int column) {
        return column == RoleColorColumn &&
               (row == InputRoleColorRow ||
                row == DropdownRoleColorRow ||
                row == SecondInputRoleColorRow ||
                row == OutputRoleColorRow);
    }

    private static bool IsDropdownRoleCell(object cellObject, CellColorRoles colorRoles) {
        var colorHex = ReadCellBackgroundHex(cellObject);

        return colorHex != null && colorRoles.DropdownColors.Contains(colorHex);
    }

    private static void ReadColumnWidths(
        object worksheetObject,
        SheetModel model,
        IProgress<SheetReadProgress>? progress) {
        dynamic worksheet = worksheetObject;
        dynamic? columns = null;

        try {
            columns = worksheet.Columns;
            var totalColumns = Math.Max(model.MaxColumn - model.FirstColumn + 1, 0);

            ReportStageProgress(
                progress,
                "Odczyt szerokości kolumn",
                5.0,
                10.0,
                0,
                totalColumns,
                totalColumns == 0 ? "Brak kolumn" : $"Kolumna 0 z {totalColumns}");

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

                var processedColumns = column - model.FirstColumn + 1;
                ReportStageProgress(
                    progress,
                    "Odczyt szerokości kolumn",
                    5.0,
                    10.0,
                    processedColumns,
                    totalColumns,
                    $"Kolumna {processedColumns} z {totalColumns}");
            }
        }
        finally {
            ReleaseComObject(columns);
        }
    }

    private static void ReadRowHeights(
        object worksheetObject,
        SheetModel model,
        IProgress<SheetReadProgress>? progress) {
        dynamic worksheet = worksheetObject;
        dynamic? rows = null;

        try {
            rows = worksheet.Rows;
            var totalRows = Math.Max(model.MaxRow - model.FirstRow + 1, 0);

            ReportStageProgress(
                progress,
                "Odczyt wysokości wierszy",
                10.0,
                15.0,
                0,
                totalRows,
                totalRows == 0 ? "Brak wierszy" : $"Wiersz 0 z {totalRows}");

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

                var processedRows = row - model.FirstRow + 1;
                ReportStageProgress(
                    progress,
                    "Odczyt wysokości wierszy",
                    10.0,
                    15.0,
                    processedRows,
                    totalRows,
                    $"Wiersz {processedRows} z {totalRows}");
            }
        }
        finally {
            ReleaseComObject(rows);
        }
    }

    private void ReadCells(
        object worksheetObject,
        SheetModel model,
        CellColorRoles colorRoles,
        IProgress<SheetReadProgress>? progress) {
        dynamic worksheet = worksheetObject;
        var mergedSlaveCells = new HashSet<(int Row, int Column)>();
        dynamic? cells = null;

        try {
            cells = worksheet.Cells;
            var values = ReadRangeValues(worksheetObject, (object)cells, model);
            var totalRows = Math.Max(model.MaxRow - model.FirstRow + 1, 0);

            ReportStageProgress(
                progress,
                "Odczyt komórek",
                15.0,
                85.0,
                0,
                totalRows,
                totalRows == 0 ? "Brak komórek" : $"Wiersz 0 z {totalRows}");

            for (var row = model.FirstRow; row <= model.MaxRow; row++) {
                var processedRows = row - model.FirstRow + 1;

                if (model.RowHeights.TryGetValue(row, out var rowHeight) && rowHeight == 0.0) {
                    ReportStageProgress(
                        progress,
                        "Odczyt komórek",
                        15.0,
                        85.0,
                        processedRows,
                        totalRows,
                        $"Wiersz {processedRows} z {totalRows}");
                    continue;
                }

                for (var column = model.FirstColumn; column <= model.MaxColumn; column++) {
                    if (model.ColumnWidths.TryGetValue(column, out var columnWidth) && columnWidth == 0.0)
                        continue;

                    if (mergedSlaveCells.Contains((row, column))) {
                        dynamic? slaveCell = null;

                        try {
                            slaveCell = cells[row, column];

                            var slaveCellModel = ReadMergedSlaveCell(
                                (object)slaveCell,
                                row,
                                column);

                            model.Cells.Add(slaveCellModel);
                        }
                        finally {
                            ReleaseComObject(slaveCell);
                        }

                        continue;
                    }

                    var rawValue = GetRangeValue(
                        values,
                        row - model.FirstRow + 1,
                        column - model.FirstColumn + 1);

                    dynamic? cell = null;

                    try {
                        cell = cells[row, column];

                        if (IsRoleColorDefinitionCell(row, column))
                            continue;

                        var cellModel = ReadRenderableCell(
                            (object)cell,
                            row,
                            column,
                            worksheetObject,
                            rawValue,
                            colorRoles);

                        if (cellModel == null)
                            continue;

                        RegisterMergedSlaveCells(cellModel, mergedSlaveCells);
                        model.Cells.Add(cellModel);
                    }
                    finally {
                        ReleaseComObject(cell);
                    }
                }

                ReportStageProgress(
                    progress,
                    "Odczyt komórek",
                    15.0,
                    85.0,
                    processedRows,
                    totalRows,
                    $"Wiersz {processedRows} z {totalRows}");
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
        object? rawValue,
        CellColorRoles colorRoles) {
        var model = new CellModel {
            Row = row,
            Column = column,
            RawValue = rawValue
        };

        if (HasBulkContent(rawValue))
            ReadDisplayText(cellObject, model);

        var hasBackground = ReadBackground(cellObject, model, colorRoles);

        ReadMerge(cellObject, model);
        ReadBorders(cellObject, model);

        if (!HasContent(model) && !hasBackground && !HasVisibleBorder(model))
            return null;

        ReadFont(cellObject, model);
        ReadAlignment(cellObject, model);

        return model;
    }


    private CellModel ReadMergedSlaveCell(
        object cellObject,
        int row,
        int column) {
        var model = new CellModel {
            Row = row,
            Column = column,
            IsMergedSlave = true
        };

        ReadBorders(cellObject, model, false);

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

    private static bool HasVisibleBorder(CellModel model) {
        return model.BorderTopThickness > 0.0
               || model.BorderBottomThickness > 0.0
               || model.BorderLeftThickness > 0.0
               || model.BorderRightThickness > 0.0;
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
        if (values is not Array valuesArray || valuesArray.Rank != 2) {
            if (rowOffset == 1 && columnOffset == 1)
                return values;

            return null;
        }

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

    private static bool ReadBackground(object cellObject, CellModel model, CellColorRoles colorRoles) {
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

            if (colorRoles.DropdownColors.Contains(colorHex)) {
                model.IsInput = true;
                model.InputType = CellInputType.ComboBox;
                return true;
            }

            if (colorRoles.InputColors.Contains(colorHex)) {
                model.IsInput = true;
                model.InputType = CellInputType.TextBox;
                return true;
            }

            if (colorRoles.OutputColors.Contains(colorHex)) {
                model.IsInput = false;
                model.InputType = CellInputType.None;
                return true;
            }

            model.IsInput = false;
            model.InputType = CellInputType.None;

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
            // ignored
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
            // ignored
        }
    }

    private void ReadBorders(object cellObject, CellModel model, bool includeMergeAreaBorders = true) {
        dynamic cell = cellObject;
        dynamic? borders = null;
        dynamic? displayFormat = null;
        dynamic? displayBorders = null;
        dynamic? mergeArea = null;
        dynamic? mergeBorders = null;
        dynamic? mergeDisplayFormat = null;
        dynamic? mergeDisplayBorders = null;

        var dominantColor = Colors.Black;

        try {
            borders = cell.Borders;
            ApplyBordersFromCollection(borders, model, ref dominantColor);
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(borders);
        }

        try {
            displayFormat = cell.DisplayFormat;
            displayBorders = displayFormat.Borders;
            ApplyBordersFromCollection(displayBorders, model, ref dominantColor);
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(displayBorders);
            ReleaseComObject(displayFormat);
        }

        try {
            if (!includeMergeAreaBorders || !Convert.ToBoolean(cell.MergeCells)) {
                model.BorderColor = dominantColor;
                return;
            }

            mergeArea = cell.MergeArea;
            mergeBorders = mergeArea.Borders;
            ApplyBordersFromCollection(mergeBorders, model, ref dominantColor);
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(mergeBorders);
        }

        try {
            if (mergeArea != null) {
                mergeDisplayFormat = mergeArea.DisplayFormat;
                mergeDisplayBorders = mergeDisplayFormat.Borders;
                ApplyBordersFromCollection(mergeDisplayBorders, model, ref dominantColor);
            }
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(mergeDisplayBorders);
            ReleaseComObject(mergeDisplayFormat);
            ReleaseComObject(mergeArea);
        }

        model.BorderColor = dominantColor;
    }

    private static void ApplyBordersFromCollection(
        dynamic borders,
        CellModel model,
        ref Color dominantColor) {
        var topColor = dominantColor;
        var bottomColor = dominantColor;
        var leftColor = dominantColor;
        var rightColor = dominantColor;

        var topThickness = ReadBorderThickness(borders, ExcelBorderEdgeTop, ref topColor);
        var bottomThickness = ReadBorderThickness(borders, ExcelBorderEdgeBottom, ref bottomColor);
        var leftThickness = ReadBorderThickness(borders, ExcelBorderEdgeLeft, ref leftColor);
        var rightThickness = ReadBorderThickness(borders, ExcelBorderEdgeRight, ref rightColor);

        if (topThickness > model.BorderTopThickness) {
            model.BorderTopThickness = topThickness;
            dominantColor = topColor;
        }

        if (bottomThickness > model.BorderBottomThickness) {
            model.BorderBottomThickness = bottomThickness;
            dominantColor = bottomColor;
        }

        if (leftThickness > model.BorderLeftThickness) {
            model.BorderLeftThickness = leftThickness;
            dominantColor = leftColor;
        }

        if (rightThickness > model.BorderRightThickness) {
            model.BorderRightThickness = rightThickness;
            dominantColor = rightColor;
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
                // ignored
            }

            try {
                return ExcelBorderWeightToDips(border.Weight);
            }
            catch {
                return 1.0;
            }
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
            // ignored
        }
        finally {
            ReleaseComObject(columns);
            ReleaseComObject(rows);
            ReleaseComObject(mergeArea);
        }
    }

    private void ReadDropdownElements(
        object worksheetObject,
        SheetModel model,
        CellColorRoles colorRoles,
        IProgress<SheetReadProgress>? progress = null,
        double overallStart = 85.0,
        double overallEnd = 92.0) {
        const int stageCount = 4;

        ReportStageProgress(
            progress,
            "Odczyt list rozwijanych",
            overallStart,
            overallEnd,
            0,
            stageCount,
            "Walidacje danych");

        ReadDataValidationDropdowns(worksheetObject, model, colorRoles);
        ReportStageProgress(progress, "Odczyt list rozwijanych", overallStart, overallEnd, 1, stageCount,
            "Kontrolki formularza");

        ReadFormControlDropdowns(worksheetObject, model, colorRoles);
        ReportStageProgress(progress, "Odczyt list rozwijanych", overallStart, overallEnd, 2, stageCount,
            "Kontrolki ActiveX");

        ReadActiveXDropdowns(worksheetObject, model, colorRoles);
        ReportStageProgress(progress, "Odczyt list rozwijanych", overallStart, overallEnd, 3, stageCount,
            "Kontrolki komórkowe");

        ReadCellControlDropdowns(worksheetObject, model);
        ReportStageProgress(progress, "Odczyt list rozwijanych", overallStart, overallEnd, 4, stageCount,
            "Gotowe");
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
                    // ignored
                }
                finally {
                    ReleaseComObject(cellControl);
                    ReleaseComObject(cell);
                }
            }
        }
        catch {
            // ignored
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
            // ignored
        }

        try {
            var source = Convert.ToString(cellControl.Source);
            var values = ReadDropdownSourceValues(source, worksheetObject);

            if (values.Count > 0)
                return values;
        }
        catch {
            // ignored
        }

        try {
            var values = ExtractValidationValues(cellControl.Items);

            if (values.Count > 0)
                return values;
        }
        catch {
            // ignored
        }

        try {
            var values = ExtractValidationValues(cellControl.Values);

            if (values.Count > 0)
                return values;
        }
        catch {
            // ignored
        }

        try {
            var values = ExtractValidationValues(cell.Validation.Formula1);

            if (values.Count > 0)
                return values;
        }
        catch {
            // ignored
        }

        return [];
    }

    private void ReadDataValidationDropdowns(object worksheetObject, SheetModel model, CellColorRoles colorRoles) {
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

                    if (!IsDropdownRoleCell((object)cell, colorRoles))
                        continue;

                    var cellModel = GetOrCreateDropdownCellModel(
                        (object)cell,
                        row,
                        column,
                        model,
                        colorRoles);

                    ReadValidation((object)cell, cellModel, worksheetObject);
                }
                finally {
                    ReleaseComObject(cell);
                }
            }
        }
        catch {
            // ignored
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
                         cell.InputType == CellInputType.ComboBox)) {
                dynamic? cell = null;

                try {
                    cell = cells[cellModel.Row, cellModel.Column];
                    ReadValidation((object)cell, cellModel, worksheetObject);
                }
                catch {
                    // ignored
                }
                finally {
                    ReleaseComObject(cell);
                }
            }
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(cells);
        }
    }

    private void ReadFormControlDropdowns(object worksheetObject, SheetModel model, CellColorRoles colorRoles) {
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

                    if (!IsDropdownRoleCell((object)topLeftCell, colorRoles))
                        continue;

                    var cellModel = GetOrCreateDropdownCellModel(
                        (object)topLeftCell,
                        row,
                        column,
                        model,
                        colorRoles);

                    var listSourceReference = ReadControlListSourceReference((object)controlFormat);
                    var linkedCellReference = ReadControlLinkedCellReference((object)controlFormat);
                    var linkedCellSelectedIndex = ReadLinkedCellSelectedIndex(linkedCellReference, worksheetObject);
                    var selectedIndex = linkedCellSelectedIndex ?? ReadFormControlSelectedIndex((object)controlFormat);
                    var selectedValue = GetDropdownValueBySelectedIndex(selectedIndex, values)
                                        ?? ReadSelectedFormControlValue((object)controlFormat, values);
                    var inputTarget = ReadLinkedCellPositionFromReference(linkedCellReference, worksheetObject);

                    ApplyDropdownValues(
                        cellModel,
                        values,
                        selectedValue,
                        inputTarget,
                        inputTarget != null,
                        selectedIndex,
                        listSourceReference,
                        linkedCellReference);

                    cellModel.DropdownControlName = Convert.ToString(shape.Name);
                    cellModel.IsActiveXDropdown = false;
                }
                catch {
                    // ignored
                }
                finally {
                    ReleaseComObject(topLeftCell);
                    ReleaseComObject(controlFormat);
                    ReleaseComObject(shape);
                }
            }
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(shapes);
        }
    }

    private void ReadActiveXDropdowns(object worksheetObject, SheetModel model, CellColorRoles colorRoles) {
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

                    if (!IsDropdownRoleCell((object)topLeftCell, colorRoles))
                        continue;

                    var cellModel = GetOrCreateDropdownCellModel(
                        (object)topLeftCell,
                        row,
                        column,
                        model,
                        colorRoles);

                    var selectedIndex = ReadActiveXSelectedIndex((object)control);
                    var selectedValue = ReadActiveXSelectedValue((object)control);
                    var listSourceReference = ReadControlListSourceReference((object)embeddedObject)
                                              ?? ReadControlListSourceReference((object)control);
                    var linkedCellReference = ReadControlLinkedCellReference((object)embeddedObject)
                                              ?? ReadControlLinkedCellReference((object)control);
                    var inputTarget = ReadLinkedCellPositionFromReference(linkedCellReference, worksheetObject);

                    ApplyDropdownValues(
                        cellModel,
                        values,
                        selectedValue,
                        inputTarget,
                        false,
                        selectedIndex,
                        listSourceReference,
                        linkedCellReference);

                    cellModel.DropdownControlName = Convert.ToString(embeddedObject.Name);
                    cellModel.IsActiveXDropdown = true;
                }
                catch {
                    // ignored
                }
                finally {
                    ReleaseComObject(topLeftCell);
                    ReleaseComObject(control);
                    ReleaseComObject(embeddedObject);
                }
            }
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(objects);
        }
    }

    private CellModel GetOrCreateDropdownCellModel(
        object cellObject,
        int row,
        int column,
        SheetModel model,
        CellColorRoles colorRoles) {
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
            // ignored
        }

        ReadDisplayText(cellObject, cellModel);
        ReadBackground(cellObject, cellModel, colorRoles);
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
                // ignored
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
            // ignored
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

        try {
            var application = worksheet.Application;

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
    }

    private static List<string> BuildEvaluationExpressions(string source) {
        var expressions = new List<string>();
        var evaluationSource = source.StartsWith('=')
            ? source
            : $"={source}";

        TryAddExpression(expressions, evaluationSource);
        TryAddExpression(expressions, NormalizeDynamicArrayFormula(evaluationSource));

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
        dynamic? application;
        object? evaluatedValue = null;

        try {
            try {
                evaluatedValue = worksheet.Evaluate(formula);
                var values = ExtractValidationValues(evaluatedValue);

                if (values.Count > 0)
                    return values;
            }
            catch {
                // ignored
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
            // ignored
        }
        finally {
            ReleaseComObject(range);
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
            // ignored
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
            // ignored
        }

        try {
            var count = Convert.ToInt32(control.ListCount);

            for (var index = 0; index < count; index++)
                AddDropdownValue(values, Convert.ToString(control.List[index]));
        }
        catch {
            // ignored
        }

        return values;
    }

    private static List<string> ReadControlFillRangeValues(
        object controlObject,
        object worksheetObject) {
        var fillRange = ReadControlListSourceReference(controlObject);

        return string.IsNullOrWhiteSpace(fillRange)
            ? []
            : ReadDropdownSourceValues(fillRange, worksheetObject);
    }

    private static string? ReadControlListSourceReference(object controlObject) {
        dynamic control = controlObject;

        try {
            var fillRange = Convert.ToString(control.ListFillRange);

            if (!string.IsNullOrWhiteSpace(fillRange))
                return fillRange;
        }
        catch {
            // ignored
        }

        try {
            var source = Convert.ToString(control.Source);

            return string.IsNullOrWhiteSpace(source)
                ? null
                : source;
        }
        catch {
            return null;
        }
    }

    private static string? ReadControlLinkedCellReference(object controlObject) {
        dynamic control = controlObject;

        try {
            var linkedCell = Convert.ToString(control.LinkedCell);

            return string.IsNullOrWhiteSpace(linkedCell)
                ? null
                : linkedCell;
        }
        catch {
            return null;
        }
    }

    private static int? ReadFormControlSelectedIndex(object controlFormatObject) {
        dynamic controlFormat = controlFormatObject;

        try {
            var selectedIndex = Convert.ToInt32(controlFormat.Value);

            return selectedIndex > 0
                ? selectedIndex
                : null;
        }
        catch {
            return null;
        }
    }

    private static int? ReadLinkedCellSelectedIndex(string? linkedCellReference, object worksheetObject) {
        if (string.IsNullOrWhiteSpace(linkedCellReference))
            return null;

        dynamic worksheet = worksheetObject;
        dynamic? workbook = null;
        dynamic? worksheets = null;
        dynamic? linkedWorksheet = null;
        dynamic? linkedCell = null;

        try {
            var reference = NormalizeExcelReference(linkedCellReference);

            if (string.IsNullOrWhiteSpace(reference))
                return null;

            var hasWorksheetReference = TrySplitWorksheetReference(
                reference,
                out var worksheetName,
                out var cellReference);

            if (hasWorksheetReference && !string.IsNullOrWhiteSpace(worksheetName)) {
                workbook = worksheet.Parent;
                worksheets = workbook.Worksheets;
                linkedWorksheet = worksheets[worksheetName];
                linkedCell = linkedWorksheet.Range[cellReference];
            }
            else
                try {
                    linkedCell = worksheet.Range[cellReference];
                }
                catch {
                    var application = worksheet.Application;
                    linkedCell = application.Range[cellReference];
                }

            return TryReadPositiveIntegerIndex(linkedCell.Value2);
        }
        catch {
            return null;
        }
        finally {
            ReleaseComObject(linkedCell);
            ReleaseComObject(linkedWorksheet);
            ReleaseComObject(worksheets);
            ReleaseComObject(workbook);
        }
    }

    private static int? TryReadPositiveIntegerIndex(object? value) {
        if (value == null)
            return null;

        if (value is int integerValue)
            return integerValue > 0 ? integerValue : null;

        if (value is double doubleValue) {
            var roundedValue = Math.Round(doubleValue);

            return doubleValue > 0 && Math.Abs(doubleValue - roundedValue) < 0.0000001
                ? Convert.ToInt32(roundedValue)
                : null;
        }

        var text = Convert.ToString(value);

        if (string.IsNullOrWhiteSpace(text))
            return null;

        if (int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var integerIndex))
            return integerIndex > 0 ? integerIndex : null;

        var normalizedText = text.Trim().Replace(',', '.');

        if (!double.TryParse(normalizedText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return null;

        var roundedNumber = Math.Round(number);

        return number > 0 && Math.Abs(number - roundedNumber) < 0.0000001
            ? Convert.ToInt32(roundedNumber)
            : null;
    }

    private static int? ReadActiveXSelectedIndex(object controlObject) {
        dynamic control = controlObject;

        try {
            var selectedIndex = Convert.ToInt32(control.ListIndex) + 1;

            return selectedIndex > 0
                ? selectedIndex
                : null;
        }
        catch {
            return null;
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

    private static string? GetDropdownValueBySelectedIndex(int? selectedIndex, IReadOnlyList<string> values) {
        if (!selectedIndex.HasValue || selectedIndex.Value <= 0 || selectedIndex.Value > values.Count)
            return null;

        return values[selectedIndex.Value - 1];
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

    private static LinkedCellPosition? ReadLinkedCellPositionFromReference(
        string? linkedCellReference,
        object worksheetObject) {
        dynamic worksheet = worksheetObject;
        dynamic? worksheets = null;
        dynamic? linkedWorksheet = null;
        dynamic? linkedCell = null;

        try {
            var reference = NormalizeExcelReference(linkedCellReference);

            if (string.IsNullOrWhiteSpace(reference))
                return null;

            var hasWorksheetReference = TrySplitWorksheetReference(
                reference,
                out var worksheetName,
                out var cellReference);

            if (hasWorksheetReference && !string.IsNullOrWhiteSpace(worksheetName)) {
                var workbook = worksheet.Parent;
                worksheets = workbook.Worksheets;
                linkedWorksheet = worksheets[worksheetName];
                linkedCell = linkedWorksheet.Range[cellReference];
            }
            else {
                try {
                    linkedCell = worksheet.Range[cellReference];
                }
                catch {
                    var application = worksheet.Application;
                    linkedCell = application.Range[cellReference];
                }

                linkedWorksheet = linkedCell.Worksheet;
            }

            return new LinkedCellPosition(
                Convert.ToString(linkedWorksheet.Name),
                Convert.ToInt32(linkedCell.Row),
                Convert.ToInt32(linkedCell.Column));
        }
        catch {
            return null;
        }
        finally {
            ReleaseComObject(linkedCell);
            ReleaseComObject(linkedWorksheet);
            ReleaseComObject(worksheets);
        }
    }

    private static string NormalizeExcelReference(string? reference) {
        if (string.IsNullOrWhiteSpace(reference))
            return string.Empty;

        var normalizedReference = reference.Trim();

        if (normalizedReference.StartsWith("=", StringComparison.Ordinal))
            normalizedReference = normalizedReference[1..].Trim();

        return normalizedReference;
    }

    private static bool TrySplitWorksheetReference(
        string reference,
        out string? worksheetName,
        out string cellReference) {
        worksheetName = null;
        cellReference = reference;

        var separatorIndex = FindWorksheetSeparator(reference);

        if (separatorIndex < 0)
            return false;

        worksheetName = NormalizeWorksheetName(reference[..separatorIndex]);
        cellReference = reference[(separatorIndex + 1)..].Trim();

        return !string.IsNullOrWhiteSpace(worksheetName) &&
               !string.IsNullOrWhiteSpace(cellReference);
    }

    private static int FindWorksheetSeparator(string reference) {
        var isInsideQuotedWorksheetName = false;

        for (var index = 0; index < reference.Length; index++) {
            if (reference[index] == '\'') {
                if (index + 1 < reference.Length && reference[index + 1] == '\'') {
                    index++;
                    continue;
                }

                isInsideQuotedWorksheetName = !isInsideQuotedWorksheetName;
                continue;
            }

            if (reference[index] == '!' && !isInsideQuotedWorksheetName)
                return index;
        }

        return -1;
    }

    private static string NormalizeWorksheetName(string worksheetNameReference) {
        var worksheetName = worksheetNameReference.Trim();

        if (worksheetName.Length >= 2 &&
            worksheetName.StartsWith("'", StringComparison.Ordinal) &&
            worksheetName.EndsWith("'", StringComparison.Ordinal))
            worksheetName = worksheetName[1..^1].Replace("''", "'");

        var workbookNameEndIndex = worksheetName.LastIndexOf(']');

        if (workbookNameEndIndex >= 0 && workbookNameEndIndex + 1 < worksheetName.Length)
            worksheetName = worksheetName[(workbookNameEndIndex + 1)..];

        return worksheetName.Trim();
    }

    private static void ApplyDropdownValues(
        CellModel model,
        IEnumerable<string> values,
        string? selectedValue,
        LinkedCellPosition? inputTarget,
        bool dropdownWritesSelectedIndex,
        int? selectedIndex = null,
        string? listSourceReference = null,
        string? linkedCellReference = null) {
        model.IsInput = true;
        model.InputType = CellInputType.ComboBox;
        model.DropdownValues = values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.CurrentCulture)
            .ToList();

        model.InputTargetRow = inputTarget?.Row;
        model.InputTargetColumn = inputTarget?.Column;
        model.InputTargetSheetName = inputTarget?.SheetName;
        model.InputTargetFormulaText = NormalizeStoredDropdownReference(linkedCellReference);
        model.DropdownWritesSelectedIndex = dropdownWritesSelectedIndex;
        model.DropdownSelectedIndex = selectedIndex > 0 ? selectedIndex : null;
        model.DropdownListSourceReference = NormalizeStoredDropdownReference(listSourceReference);
        model.DropdownLinkedCellReference = NormalizeStoredDropdownReference(linkedCellReference);

        if (!string.IsNullOrWhiteSpace(selectedValue))
            model.DisplayText = selectedValue;
    }

    private static string? NormalizeStoredDropdownReference(string? reference) {
        if (string.IsNullOrWhiteSpace(reference))
            return null;

        var normalizedReference = reference.Trim();

        return string.IsNullOrWhiteSpace(normalizedReference)
            ? null
            : normalizedReference;
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

    private static void ReadMacroButtons(
        object worksheetObject,
        SheetModel model,
        IProgress<SheetReadProgress>? progress = null,
        double overallStart = 97.0,
        double overallEnd = 100.0) {
        model.MacroButtons.Clear();

        ReportStageProgress(progress, "Odczyt przycisków i makr", overallStart, overallEnd, 0, 2,
            "Przyciski arkusza");
        ReadShapeMacroButtons(worksheetObject, model);

        ReportStageProgress(progress, "Odczyt przycisków i makr", overallStart, overallEnd, 1, 2,
            "Przyciski ActiveX");
        ReadActiveXCommandButtons(worksheetObject, model);

        ReportStageProgress(progress, "Odczyt przycisków i makr", overallStart, overallEnd, 2, 2,
            "Gotowe");
    }

    private static void ReadShapeMacroButtons(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic? shapes = null;
        dynamic? cells = null;
        dynamic? originCell = null;

        try {
            shapes = worksheet.Shapes;
            var shapeCount = Convert.ToInt32(shapes.Count);

            if (shapeCount == 0)
                return;

            cells = worksheet.Cells;
            originCell = cells[model.FirstRow, model.FirstColumn];

            var originLeft = Convert.ToDouble(originCell.Left) * PointsToDips;
            var originTop = Convert.ToDouble(originCell.Top) * PointsToDips;

            for (var index = 1; index <= shapeCount; index++) {
                dynamic? shape = null;

                try {
                    shape = shapes.Item(index);

                    if (!Convert.ToBoolean(shape.Visible))
                        continue;

                    if (IsDropdownFormControl((object)shape))
                        continue;

                    var macroName = NormalizeMacroName(Convert.ToString(shape.OnAction));

                    if (string.IsNullOrWhiteSpace(macroName))
                        continue;

                    var left = Convert.ToDouble(shape.Left) * PointsToDips - originLeft;
                    var top = Convert.ToDouble(shape.Top) * PointsToDips - originTop;
                    var width = Math.Max(Convert.ToDouble(shape.Width) * PointsToDips, 1.0);
                    var height = Math.Max(Convert.ToDouble(shape.Height) * PointsToDips, 1.0);

                    if (left + width < 0.0 || top + height < 0.0)
                        continue;

                    model.MacroButtons.Add(new MacroButtonConfig {
                        Label = ReadShapeText((object)shape),
                        MacroName = macroName,
                        Tooltip = $"Uruchamia makro: {macroName}",
                        ShapeName = Convert.ToString(shape.Name) ?? string.Empty,
                        Left = left,
                        Top = top,
                        Width = width,
                        Height = height,
                        ZIndex = index,
                        IsSheetButton = true
                    });
                }
                catch {
                    // ignored
                }
                finally {
                    ReleaseComObject(shape);
                }
            }
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(originCell);
            ReleaseComObject(cells);
            ReleaseComObject(shapes);
        }
    }

    private static void ReadActiveXCommandButtons(object worksheetObject, SheetModel model) {
        dynamic worksheet = worksheetObject;
        dynamic? objects = null;
        dynamic? cells = null;
        dynamic? originCell = null;

        try {
            objects = worksheet.OLEObjects();
            var objectCount = Convert.ToInt32(objects.Count);

            if (objectCount == 0)
                return;

            cells = worksheet.Cells;
            originCell = cells[model.FirstRow, model.FirstColumn];

            var originLeft = Convert.ToDouble(originCell.Left) * PointsToDips;
            var originTop = Convert.ToDouble(originCell.Top) * PointsToDips;

            for (var index = 1; index <= objectCount; index++) {
                dynamic? embeddedObject = null;
                dynamic? control = null;

                try {
                    embeddedObject = objects.Item(index);

                    if (!IsActiveXCommandButton((object)embeddedObject))
                        continue;

                    control = embeddedObject.Object;

                    var name = Convert.ToString(embeddedObject.Name) ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    var left = Convert.ToDouble(embeddedObject.Left) * PointsToDips - originLeft;
                    var top = Convert.ToDouble(embeddedObject.Top) * PointsToDips - originTop;
                    var width = Math.Max(Convert.ToDouble(embeddedObject.Width) * PointsToDips, 1.0);
                    var height = Math.Max(Convert.ToDouble(embeddedObject.Height) * PointsToDips, 1.0);

                    if (left + width < 0.0 || top + height < 0.0)
                        continue;

                    var label = ReadActiveXCaption((object)control);
                    var displayLabel = string.IsNullOrWhiteSpace(label)
                        ? name
                        : label;

                    model.MacroButtons.Add(new MacroButtonConfig {
                        Label = displayLabel,
                        MacroName = name,
                        Tooltip = $"Uruchamia przycisk ActiveX: {name}",
                        OleObjectName = name,
                        Left = left,
                        Top = top,
                        Width = width,
                        Height = height,
                        ZIndex = 5000 + index,
                        IsSheetButton = true,
                        IsActiveXCommandButton = true,
                        RefreshLayoutAfterRun = true
                    });
                }
                catch {
                    // ignored
                }
                finally {
                    ReleaseComObject(control);
                    ReleaseComObject(embeddedObject);
                }
            }
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(originCell);
            ReleaseComObject(cells);
            ReleaseComObject(objects);
        }
    }

    private static bool IsDropdownFormControl(object shapeObject) {
        dynamic shape = shapeObject;

        try {
            if (Convert.ToInt32(shape.Type) != ExcelShapeTypeFormControl)
                return false;

            var controlType = Convert.ToInt32(shape.FormControlType);

            return controlType == ExcelFormControlTypeDropdown ||
                   controlType == ExcelFormControlTypeListBox;
        }
        catch {
            return false;
        }
    }

    private static bool IsActiveXCommandButton(object embeddedObjectObject) {
        dynamic embeddedObject = embeddedObjectObject;

        try {
            var progId = Convert.ToString(embeddedObject.ProgId)
                         ?? Convert.ToString(embeddedObject.progID)
                         ?? string.Empty;

            if (progId.Contains("CommandButton", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        catch {
            // ignored
        }

        try {
            var control = embeddedObject.Object;
            var caption = Convert.ToString(control.Caption);
            ReleaseComObject(control);

            return !string.IsNullOrWhiteSpace(caption);
        }
        catch {
            return false;
        }
    }

    private static string ReadShapeText(object shapeObject) {
        dynamic shape = shapeObject;

        try {
            var text = Convert.ToString(shape.TextFrame2.TextRange.Text);

            if (!string.IsNullOrWhiteSpace(text))
                return text.Trim();
        }
        catch {
            // ignored
        }

        try {
            var text = Convert.ToString(shape.TextFrame.Characters().Text);

            if (!string.IsNullOrWhiteSpace(text))
                return text.Trim();
        }
        catch {
            // ignored
        }

        try {
            var text = Convert.ToString(shape.Name);

            return text?.Trim() ?? string.Empty;
        }
        catch {
            return string.Empty;
        }
    }

    private static string ReadActiveXCaption(object controlObject) {
        dynamic control = controlObject;

        try {
            var caption = Convert.ToString(control.Caption);

            return caption?.Trim() ?? string.Empty;
        }
        catch {
            return string.Empty;
        }
    }

    private static string NormalizeMacroName(string? macroName) {
        if (string.IsNullOrWhiteSpace(macroName))
            return string.Empty;

        return macroName.Trim();
    }

    private static void ReadImages(
        object worksheetObject,
        SheetModel model,
        IProgress<SheetReadProgress>? progress = null,
        double overallStart = 92.0,
        double overallEnd = 97.0) {
        dynamic worksheet = worksheetObject;
        dynamic? shapes = null;
        dynamic? cells = null;
        dynamic? originCell = null;

        model.Images.Clear();

        ReportProgress(progress, "Odczyt obrazów", overallStart, 0.0, "Sprawdzanie elementów graficznych…");

        try {
            shapes = worksheet.Shapes;

            var shapeCount = Convert.ToInt32(shapes.Count);

            if (shapeCount == 0) {
                ReportProgress(progress, "Odczyt obrazów", overallEnd, 100.0, "Brak obrazów");
                return;
            }

            cells = worksheet.Cells;
            originCell = cells[model.FirstRow, model.FirstColumn];

            var originLeft = Convert.ToDouble(originCell.Left) * PointsToDips;
            var originTop = Convert.ToDouble(originCell.Top) * PointsToDips;

            for (var index = 1; index <= shapeCount; index++) {
                dynamic? shape = null;

                try {
                    shape = shapes.Item(index);

                    if (!IsRenderableSheetVisualShape((object)shape))
                        continue;

                    var imageBytes = TryCopyShapeAsPngFromClipboard((object)shape);

                    if (imageBytes == null || imageBytes.Length == 0)
                        continue;

                    model.Images.Add(new SheetImageModel {
                        Name = Convert.ToString(shape.Name) ?? string.Empty,
                        ImageBytes = imageBytes,
                        Left = Convert.ToDouble(shape.Left) * PointsToDips - originLeft,
                        Top = Convert.ToDouble(shape.Top) * PointsToDips - originTop,
                        Width = Math.Max(Convert.ToDouble(shape.Width) * PointsToDips, 1.0),
                        Height = Math.Max(Convert.ToDouble(shape.Height) * PointsToDips, 1.0),
                        ZIndex = index
                    });
                }
                catch {
                    // ignored
                }
                finally {
                    ReleaseComObject(shape);
                }
            }
        }
        catch {
            // ignored
        }
        finally {
            ReleaseComObject(originCell);
            ReleaseComObject(cells);
            ReleaseComObject(shapes);
            ReportProgress(progress, "Odczyt obrazów", overallEnd, 100.0, "Gotowe");
        }
    }

    private static bool IsRenderableSheetVisualShape(object shapeObject) {
        dynamic shape = shapeObject;

        try {
            if (!Convert.ToBoolean(shape.Visible))
                return false;
        }
        catch {
            // ignored
        }

        try {
            var shapeType = Convert.ToInt32(shape.Type);

            return shapeType is ExcelShapeTypeLinkedPicture
                or ExcelShapeTypePicture
                or ExcelShapeTypeGraphic
                or ExcelShapeTypeLinkedGraphic
                or ExcelShapeTypeLine;
        }
        catch {
            return false;
        }
    }

    private static byte[]? TryCopyShapeAsPngFromClipboard(object shapeObject) {
        dynamic shape = shapeObject;

        try {
            shape.CopyPicture(
                ExcelCopyPictureAppearanceScreen,
                ExcelCopyPictureFormatBitmap);

            for (var attempt = 0; attempt < MaximumClipboardReadAttempts; attempt++) {
                try {
                    var bitmap = Clipboard.GetImage();

                    if (bitmap != null) {
                        var encoder = new PngBitmapEncoder();
                        encoder.Frames.Add(BitmapFrame.Create(bitmap));

                        using var stream = new MemoryStream();
                        encoder.Save(stream);
                        return stream.ToArray();
                    }
                }
                catch {
                    // ignored
                }

                Thread.Sleep(ClipboardRetryDelayMilliseconds);
            }
        }
        catch {
            // ignored
        }

        return null;
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

    private static void ReportStageProgress(
        IProgress<SheetReadProgress>? progress,
        string stage,
        double overallStart,
        double overallEnd,
        int current,
        int total,
        string detail) {
        if (progress == null)
            return;

        var stagePercentage = total <= 0
            ? 100.0
            : Math.Clamp(current * 100.0 / total, 0.0, 100.0);

        var overallPercentage = overallStart +
                                (overallEnd - overallStart) * stagePercentage / 100.0;

        ReportProgress(
            progress,
            stage,
            overallPercentage,
            stagePercentage,
            detail);
    }

    private static void ReportProgress(
        IProgress<SheetReadProgress>? progress,
        string stage,
        double overallPercentage,
        double stagePercentage,
        string detail) {
        progress?.Report(new SheetReadProgress(
            stage,
            Math.Clamp(overallPercentage, 0.0, 100.0),
            Math.Clamp(stagePercentage, 0.0, 100.0),
            detail));
    }

    private static void ReleaseComObject(object? comObject) {
        if (comObject == null)
            return;

        try {
            if (Marshal.IsComObject(comObject))
                Marshal.FinalReleaseComObject(comObject);
        }
        catch (InvalidComObjectException) {
            // The RCW was already released through another reference.
        }
        catch (COMException) {
            // Excel may already be shutting down.
        }
    }

    private sealed record LinkedCellPosition(
        string? SheetName,
        int Row,
        int Column);

    private sealed class CellColorRoles {
        public HashSet<string> InputColors { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> DropdownColors { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> OutputColors { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private sealed record ReadBounds(
        int FirstRow,
        int FirstColumn,
        int MaximumRow,
        int MaximumColumn);
}