using System.Globalization;
using System.IO;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public class ExcelSessionService : IDisposable {
    private const int ExcelCalculationManual = -4135;

    private dynamic? _application;
    private bool _disposed;
    private dynamic? _workbook;

    public void Dispose() {
        if (_disposed) return;

        _disposed = true;
        CloseWorkbookInternal();

        if (_application != null)
            try {
                _application.Quit();
            }
            catch {
            }
            finally {
                ReleaseComObject(_application);
                _application = null;
            }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    public void OpenSession(string workbookPath) {
        EnsureApplicationCreated();
        CloseWorkbookInternal();
        OpenWorkbookInternal(workbookPath);
    }

    public dynamic GetFirstWorksheet() {
        if (_workbook == null)
            throw new InvalidOperationException("Brak otwartego skoroszytu.");

        dynamic? worksheets = null;

        try {
            worksheets = _workbook.Worksheets;

            if (Convert.ToInt32(worksheets.Count) == 0)
                throw new InvalidOperationException("Skoroszyt nie zawiera żadnego arkusza.");

            return worksheets[1];
        }
        finally {
            ReleaseComObject(worksheets);
        }
    }

    private dynamic GetWorksheet(string? worksheetName) {
        if (string.IsNullOrWhiteSpace(worksheetName))
            return GetFirstWorksheet();

        if (_workbook == null)
            throw new InvalidOperationException("Brak otwartego skoroszytu.");

        dynamic? worksheets = null;

        try {
            worksheets = _workbook.Worksheets;
            return worksheets[worksheetName];
        }
        finally {
            ReleaseComObject(worksheets);
        }
    }

    public void ExecuteWithEventsEnabled(Action action) {
        if (action == null)
            throw new ArgumentNullException(nameof(action));

        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        bool? previousEnableEvents = null;

        try {
            previousEnableEvents = Convert.ToBoolean(_application.EnableEvents);
            _application.EnableEvents = true;
        }
        catch {
        }

        try {
            action();
        }
        finally {
            if (previousEnableEvents.HasValue)
                try {
                    _application.EnableEvents = previousEnableEvents.Value;
                }
                catch {
                }
        }
    }

    public void WriteCellValue(int row, int column, object? value) {
        WriteCellValue(row, column, value, null);
    }

    public void WriteCellValue(int row, int column, object? value, string? worksheetName) {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        dynamic? worksheet = null;
        dynamic? cells = null;
        dynamic? cell = null;

        try {
            worksheet = GetWorksheet(worksheetName);
            cells = worksheet.Cells;
            cell = cells[row, column];
            cell.Value2 = value;
        }
        finally {
            ReleaseComObject(cell);
            ReleaseComObject(cells);
            ReleaseComObject(worksheet);
        }
    }

    public bool TryWriteDropdownControlValue(CellModel cellModel, string selectedValue, int selectedIndex) {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        var wasWritten = cellModel.IsActiveXDropdown
            ? TryWriteActiveXDropdownControlValue(cellModel, selectedValue, selectedIndex)
            : TryWriteFormControlDropdownValue(cellModel, selectedValue, selectedIndex);

        if (!wasWritten)
            wasWritten = cellModel.IsActiveXDropdown
                ? TryWriteFormControlDropdownValue(cellModel, selectedValue, selectedIndex)
                : TryWriteActiveXDropdownControlValue(cellModel, selectedValue, selectedIndex);

        var linkedCellWasWritten = TryWriteDropdownLinkedCellValue(cellModel, selectedValue, selectedIndex);

        return wasWritten || linkedCellWasWritten;
    }

    private bool TryWriteFormControlDropdownValue(CellModel cellModel, string selectedValue, int selectedIndex) {
        dynamic? worksheets = null;

        try {
            worksheets = _workbook?.Worksheets;

            if (worksheets == null)
                return false;

            var worksheetCount = Convert.ToInt32(worksheets.Count);

            for (var worksheetIndex = 1; worksheetIndex <= worksheetCount; worksheetIndex++) {
                dynamic? worksheet = null;

                try {
                    worksheet = worksheets[worksheetIndex];

                    if (TryWriteNamedFormControlDropdownValue((object)worksheet, cellModel, selectedValue,
                            selectedIndex))
                        return true;

                    if (TryWritePositionedFormControlDropdownValue((object)worksheet, cellModel, selectedValue,
                            selectedIndex))
                        return true;
                }
                catch {
                }
                finally {
                    ReleaseComObject(worksheet);
                }
            }

            return false;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(worksheets);
        }
    }

    private static bool TryWriteNamedFormControlDropdownValue(
        object worksheetObject,
        CellModel cellModel,
        string selectedValue,
        int selectedIndex) {
        if (string.IsNullOrWhiteSpace(cellModel.DropdownControlName))
            return false;

        dynamic worksheet = worksheetObject;
        dynamic? shapes = null;
        dynamic? shape = null;

        try {
            shapes = worksheet.Shapes;
            shape = shapes.Item(cellModel.DropdownControlName);

            if (!IsDropdownFormControlShape((object)shape))
                return false;

            return TryWriteFormControlDropdownShape((object)shape, cellModel, selectedValue, selectedIndex);
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(shape);
            ReleaseComObject(shapes);
        }
    }

    private static bool TryWritePositionedFormControlDropdownValue(
        object worksheetObject,
        CellModel cellModel,
        string selectedValue,
        int selectedIndex) {
        dynamic worksheet = worksheetObject;
        dynamic? shapes = null;

        try {
            shapes = worksheet.Shapes;
            var shapeCount = Convert.ToInt32(shapes.Count);

            for (var shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex++) {
                dynamic? shape = null;

                try {
                    shape = shapes.Item(shapeIndex);

                    if (!IsDropdownFormControlShape((object)shape))
                        continue;

                    if (!IsControlPlacedOnCell((object)shape, cellModel.Row, cellModel.Column) &&
                        !IsControlLinkedToCell((object)shape, cellModel))
                        continue;

                    if (TryWriteFormControlDropdownShape((object)shape, cellModel, selectedValue, selectedIndex))
                        return true;
                }
                catch {
                }
                finally {
                    ReleaseComObject(shape);
                }
            }

            return false;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(shapes);
        }
    }

    private static bool TryWriteFormControlDropdownShape(
        object shapeObject,
        CellModel cellModel,
        string selectedValue,
        int selectedIndex) {
        dynamic shape = shapeObject;
        dynamic? controlFormat = null;

        try {
            controlFormat = shape.ControlFormat;

            var resolvedSelectedIndex = FindFormControlSelectedIndex(
                (object)controlFormat,
                selectedValue,
                selectedIndex);

            if (resolvedSelectedIndex <= 0)
                return false;

            controlFormat.Value = resolvedSelectedIndex;
            TryWriteControlLinkedCell((object)controlFormat, cellModel, resolvedSelectedIndex);

            return true;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(controlFormat);
        }
    }

    private static bool IsDropdownFormControlShape(object shapeObject) {
        dynamic shape = shapeObject;

        try {
            const int excelShapeTypeFormControl = 8;
            const int excelFormControlTypeDropdown = 2;
            const int excelFormControlTypeListBox = 6;

            if (Convert.ToInt32(shape.Type) != excelShapeTypeFormControl)
                return false;

            var controlType = Convert.ToInt32(shape.FormControlType);

            return controlType == excelFormControlTypeDropdown ||
                   controlType == excelFormControlTypeListBox;
        }
        catch {
            return false;
        }
    }

    private static int FindFormControlSelectedIndex(
        object controlFormatObject,
        string selectedValue,
        int fallbackSelectedIndex) {
        dynamic controlFormat = controlFormatObject;

        if (string.IsNullOrWhiteSpace(selectedValue))
            return fallbackSelectedIndex;

        var normalizedSelectedValue = NormalizeDropdownText(selectedValue);

        try {
            var count = Convert.ToInt32(controlFormat.ListCount);

            for (var index = 1; index <= count; index++)
                if (AreDropdownValuesEqual(Convert.ToString(controlFormat.List[index]), selectedValue,
                        normalizedSelectedValue))
                    return index;
        }
        catch {
        }

        if (fallbackSelectedIndex > 0)
            return fallbackSelectedIndex;

        return int.TryParse(selectedValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture,
                   out var numericIndex)
               && numericIndex > 0
            ? numericIndex
            : 0;
    }

    private bool TryWriteActiveXDropdownControlValue(CellModel cellModel, string selectedValue, int selectedIndex) {
        dynamic? worksheets = null;

        try {
            worksheets = _workbook?.Worksheets;

            if (worksheets == null)
                return false;

            var worksheetCount = Convert.ToInt32(worksheets.Count);

            for (var worksheetIndex = 1; worksheetIndex <= worksheetCount; worksheetIndex++) {
                dynamic? worksheet = null;

                try {
                    worksheet = worksheets[worksheetIndex];

                    if (TryWriteNamedActiveXDropdownControlValue((object)worksheet, cellModel, selectedValue,
                            selectedIndex))
                        return true;

                    if (TryWritePositionedActiveXDropdownControlValue((object)worksheet, cellModel, selectedValue,
                            selectedIndex))
                        return true;
                }
                catch {
                }
                finally {
                    ReleaseComObject(worksheet);
                }
            }

            return false;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(worksheets);
        }
    }

    private static bool TryWriteNamedActiveXDropdownControlValue(
        object worksheetObject,
        CellModel cellModel,
        string selectedValue,
        int selectedIndex) {
        if (string.IsNullOrWhiteSpace(cellModel.DropdownControlName))
            return false;

        dynamic worksheet = worksheetObject;
        dynamic? objects = null;
        dynamic? embeddedObject = null;

        try {
            objects = worksheet.OLEObjects();
            embeddedObject = objects.Item(cellModel.DropdownControlName);

            if (!IsActiveXDropdownObject((object)embeddedObject))
                return false;

            return TryWriteActiveXDropdownObject((object)embeddedObject, cellModel, selectedValue, selectedIndex);
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(embeddedObject);
            ReleaseComObject(objects);
        }
    }

    private static bool TryWritePositionedActiveXDropdownControlValue(
        object worksheetObject,
        CellModel cellModel,
        string selectedValue,
        int selectedIndex) {
        dynamic worksheet = worksheetObject;
        dynamic? objects = null;

        try {
            objects = worksheet.OLEObjects();
            var objectCount = Convert.ToInt32(objects.Count);

            for (var objectIndex = 1; objectIndex <= objectCount; objectIndex++) {
                dynamic? embeddedObject = null;

                try {
                    embeddedObject = objects.Item(objectIndex);

                    if (!IsActiveXDropdownObject((object)embeddedObject))
                        continue;

                    if (!IsControlPlacedOnCell((object)embeddedObject, cellModel.Row, cellModel.Column) &&
                        !IsControlLinkedToCell((object)embeddedObject, cellModel))
                        continue;

                    if (TryWriteActiveXDropdownObject((object)embeddedObject, cellModel, selectedValue, selectedIndex))
                        return true;
                }
                catch {
                }
                finally {
                    ReleaseComObject(embeddedObject);
                }
            }

            return false;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(objects);
        }
    }

    private static bool TryWriteActiveXDropdownObject(
        object embeddedObjectObject,
        CellModel cellModel,
        string selectedValue,
        int selectedIndex) {
        dynamic embeddedObject = embeddedObjectObject;
        dynamic? control = null;

        try {
            control = embeddedObject.Object;

            var wasWritten = false;
            var resolvedSelectedIndex = FindActiveXSelectedIndex((object)control, selectedValue, selectedIndex);

            if (resolvedSelectedIndex > 0)
                try {
                    control.ListIndex = resolvedSelectedIndex - 1;
                    wasWritten = true;
                }
                catch {
                }

            try {
                control.Value = selectedValue;
                wasWritten = true;
            }
            catch {
            }

            if (wasWritten) {
                TryWriteControlLinkedCell((object)embeddedObject, cellModel, selectedValue);

                if (control != null)
                    TryWriteControlLinkedCell((object)control, cellModel, selectedValue);
            }

            return wasWritten;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(control);
        }
    }

    private static bool IsActiveXDropdownObject(object embeddedObjectObject) {
        dynamic embeddedObject = embeddedObjectObject;
        dynamic? control = null;

        try {
            var progId = Convert.ToString(embeddedObject.ProgId)
                         ?? Convert.ToString(embeddedObject.progID)
                         ?? string.Empty;

            if (progId.Contains("ComboBox", StringComparison.OrdinalIgnoreCase) ||
                progId.Contains("ListBox", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        catch {
        }

        try {
            control = embeddedObject.Object;
            _ = control.ListCount;

            return true;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(control);
        }
    }

    private static int FindActiveXSelectedIndex(
        object controlObject,
        string selectedValue,
        int fallbackSelectedIndex) {
        dynamic control = controlObject;

        if (string.IsNullOrWhiteSpace(selectedValue))
            return fallbackSelectedIndex;

        var normalizedSelectedValue = NormalizeDropdownText(selectedValue);

        try {
            var count = Convert.ToInt32(control.ListCount);

            for (var index = 0; index < count; index++)
                if (AreDropdownValuesEqual(Convert.ToString(control.List[index]), selectedValue,
                        normalizedSelectedValue))
                    return index + 1;
        }
        catch {
        }

        if (fallbackSelectedIndex > 0)
            return fallbackSelectedIndex;

        return int.TryParse(selectedValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture,
                   out var numericIndex)
               && numericIndex > 0
            ? numericIndex
            : 0;
    }

    private static bool IsControlPlacedOnCell(object controlObject, int row, int column) {
        dynamic control = controlObject;
        dynamic? topLeftCell = null;
        dynamic? bottomRightCell = null;

        try {
            topLeftCell = control.TopLeftCell;
            bottomRightCell = control.BottomRightCell;

            var topRow = Convert.ToInt32(topLeftCell.Row);
            var leftColumn = Convert.ToInt32(topLeftCell.Column);
            var bottomRow = Convert.ToInt32(bottomRightCell.Row);
            var rightColumn = Convert.ToInt32(bottomRightCell.Column);

            return row >= topRow &&
                   row <= bottomRow &&
                   column >= leftColumn &&
                   column <= rightColumn;
        }
        catch {
            try {
                if (topLeftCell == null)
                    topLeftCell = control.TopLeftCell;

                return Convert.ToInt32(topLeftCell.Row) == row &&
                       Convert.ToInt32(topLeftCell.Column) == column;
            }
            catch {
                return false;
            }
        }
        finally {
            ReleaseComObject(bottomRightCell);
            ReleaseComObject(topLeftCell);
        }
    }

    private static bool IsControlLinkedToCell(object controlObject, CellModel cellModel) {
        if (cellModel is not { InputTargetRow: not null, InputTargetColumn: not null })
            return false;

        dynamic control = controlObject;

        try {
            var linkedCellReference = Convert.ToString(control.LinkedCell);

            if (string.IsNullOrWhiteSpace(linkedCellReference))
                return false;

            return IsReferenceMatchingCell(
                linkedCellReference,
                cellModel.InputTargetSheetName,
                cellModel.InputTargetRow.Value,
                cellModel.InputTargetColumn.Value);
        }
        catch {
            try {
                var controlFormat = control.ControlFormat;
                var linkedCellReference = Convert.ToString(controlFormat.LinkedCell);

                return IsReferenceMatchingCell(
                    linkedCellReference,
                    cellModel.InputTargetSheetName,
                    cellModel.InputTargetRow.Value,
                    cellModel.InputTargetColumn.Value);
            }
            catch {
                return false;
            }
        }
    }

    private static bool IsReferenceMatchingCell(string? reference, string? sheetName, int row, int column) {
        if (string.IsNullOrWhiteSpace(reference))
            return false;

        var normalizedReference = reference.Trim();

        if (normalizedReference.StartsWith('='))
            normalizedReference = normalizedReference[1..].Trim();

        string? referenceSheetName = null;
        var cellAddress = normalizedReference;
        var separatorIndex = normalizedReference.LastIndexOf('!');

        if (separatorIndex >= 0 && separatorIndex < normalizedReference.Length - 1) {
            referenceSheetName = NormalizeSheetNameFromReference(normalizedReference[..separatorIndex]);
            cellAddress = normalizedReference[(separatorIndex + 1)..].Trim();
        }

        if (!string.IsNullOrWhiteSpace(referenceSheetName) &&
            !string.IsNullOrWhiteSpace(sheetName) &&
            !string.Equals(referenceSheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            return false;

        return TryParseCellAddress(cellAddress, out var referenceRow, out var referenceColumn) &&
               referenceRow == row &&
               referenceColumn == column;
    }

    private static bool TryParseCellAddress(string address, out int row, out int column) {
        row = 0;
        column = 0;

        if (string.IsNullOrWhiteSpace(address))
            return false;

        var normalizedAddress = address.Trim().Replace("$", string.Empty);
        var columnLetters = string.Empty;
        var rowDigits = string.Empty;

        foreach (var character in normalizedAddress) {
            if (char.IsLetter(character) && rowDigits.Length == 0) {
                columnLetters += char.ToUpperInvariant(character);
                continue;
            }

            if (char.IsDigit(character)) {
                rowDigits += character;
                continue;
            }

            return false;
        }

        if (columnLetters.Length == 0 || rowDigits.Length == 0)
            return false;

        foreach (var character in columnLetters)
            column = column * 26 + character - 'A' + 1;

        return int.TryParse(rowDigits, NumberStyles.Integer, CultureInfo.InvariantCulture, out row);
    }

    private static bool TrySplitSheetReference(
        string reference,
        out string? sheetName,
        out string cellAddress) {
        sheetName = null;
        cellAddress = string.Empty;

        if (string.IsNullOrWhiteSpace(reference))
            return false;

        var normalizedReference = reference.Trim();

        if (normalizedReference.StartsWith('='))
            normalizedReference = normalizedReference[1..].Trim();

        var separatorIndex = normalizedReference.LastIndexOf('!');

        if (separatorIndex <= 0 || separatorIndex >= normalizedReference.Length - 1)
            return false;

        sheetName = NormalizeSheetNameFromReference(normalizedReference[..separatorIndex]);
        cellAddress = normalizedReference[(separatorIndex + 1)..].Trim();

        return !string.IsNullOrWhiteSpace(sheetName) &&
               !string.IsNullOrWhiteSpace(cellAddress);
    }

    private static string NormalizeSheetNameFromReference(string sheetReference) {
        var sheetName = sheetReference.Trim();

        if (sheetName.StartsWith('\'') && sheetName.EndsWith('\'') && sheetName.Length >= 2)
            sheetName = sheetName[1..^1].Replace("''", "'");

        var workbookSeparatorIndex = sheetName.LastIndexOf(']');

        if (workbookSeparatorIndex >= 0 && workbookSeparatorIndex < sheetName.Length - 1)
            sheetName = sheetName[(workbookSeparatorIndex + 1)..];

        return sheetName.Trim();
    }

    private static void TryWriteControlLinkedCell(object controlObject, CellModel cellModel, object? value) {
        if (TryGetControlLinkedCellReference(controlObject, out var linkedCellReference) &&
            TryWriteControlLinkedCellReference(controlObject, linkedCellReference, value))
            return;

        TryWriteControlInputTargetCell(controlObject, cellModel, value);
    }

    private static bool TryGetControlLinkedCellReference(object controlObject, out string linkedCellReference) {
        linkedCellReference = string.Empty;
        dynamic control = controlObject;

        try {
            linkedCellReference = Convert.ToString(control.LinkedCell) ?? string.Empty;
        }
        catch {
        }

        if (!string.IsNullOrWhiteSpace(linkedCellReference))
            return true;

        try {
            var controlFormat = control.ControlFormat;
            linkedCellReference = Convert.ToString(controlFormat.LinkedCell) ?? string.Empty;
        }
        catch {
        }

        return !string.IsNullOrWhiteSpace(linkedCellReference);
    }

    private static bool TryWriteControlLinkedCellReference(
        object controlObject,
        string linkedCellReference,
        object? value) {
        var normalizedReference = linkedCellReference.Trim();

        if (normalizedReference.StartsWith('='))
            normalizedReference = normalizedReference[1..].Trim();

        if (string.IsNullOrWhiteSpace(normalizedReference))
            return false;

        dynamic? worksheet = null;
        dynamic? targetWorksheet = null;
        dynamic? workbook = null;
        dynamic? range = null;

        try {
            worksheet = GetControlWorksheet(controlObject);

            if (worksheet == null)
                return false;

            if (TrySplitSheetReference(normalizedReference, out var sheetName, out var cellAddress) &&
                !string.IsNullOrWhiteSpace(sheetName) &&
                !string.IsNullOrWhiteSpace(cellAddress)) {
                workbook = worksheet.Parent;
                targetWorksheet = workbook.Worksheets[sheetName];
                range = targetWorksheet.Range[cellAddress];
            }
            else
                range = worksheet.Range[normalizedReference];

            range.Value2 = value;
            return true;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(range);
            ReleaseComObject(targetWorksheet);
            ReleaseComObject(workbook);
            ReleaseComObject(worksheet);
        }
    }

    private static bool TryWriteControlInputTargetCell(object controlObject, CellModel cellModel, object? value) {
        if (cellModel is not { InputTargetRow: not null, InputTargetColumn: not null })
            return false;

        dynamic? worksheet = null;
        dynamic? workbook = null;
        dynamic? targetWorksheet = null;
        dynamic? cells = null;
        dynamic? cell = null;

        try {
            worksheet = GetControlWorksheet(controlObject);

            if (worksheet == null)
                return false;

            targetWorksheet = worksheet;

            if (!string.IsNullOrWhiteSpace(cellModel.InputTargetSheetName)) {
                workbook = worksheet.Parent;
                targetWorksheet = workbook.Worksheets[cellModel.InputTargetSheetName];
            }

            cells = targetWorksheet.Cells;
            cell = cells[cellModel.InputTargetRow.Value, cellModel.InputTargetColumn.Value];
            cell.Value2 = value;

            return true;
        }
        catch {
            return false;
        }
        finally {
            ReleaseComObject(cell);
            ReleaseComObject(cells);

            if (!ReferenceEquals(targetWorksheet, worksheet))
                ReleaseComObject(targetWorksheet);

            ReleaseComObject(workbook);
            ReleaseComObject(worksheet);
        }
    }

    private static dynamic? GetControlWorksheet(object controlObject) {
        dynamic control = controlObject;
        dynamic? parent = null;

        try {
            parent = control.Parent;

            try {
                _ = parent.Cells;
                return parent;
            }
            catch {
            }

            try {
                var parentParent = parent.Parent;
                _ = parentParent.Cells;
                return parentParent;
            }
            catch {
                return null;
            }
        }
        catch {
            return null;
        }
    }

    private bool TryWriteDropdownLinkedCellValue(
        CellModel cellModel,
        string selectedValue,
        int selectedIndex) {
        if (cellModel is not { InputTargetRow: not null, InputTargetColumn: not null })
            return false;

        if (!TryCreateDropdownLinkedCellValue(cellModel, selectedValue, selectedIndex, out var linkedCellValue))
            return false;

        try {
            WriteCellValue(
                cellModel.InputTargetRow.Value,
                cellModel.InputTargetColumn.Value,
                linkedCellValue,
                cellModel.InputTargetSheetName);

            return true;
        }
        catch {
            return false;
        }
    }

    private static bool TryCreateDropdownLinkedCellValue(
        CellModel cellModel,
        string selectedValue,
        int selectedIndex,
        out object? value) {
        if (cellModel.DropdownWritesSelectedIndex) {
            if (selectedIndex > 0) {
                value = selectedIndex;
                return true;
            }

            value = null;
            return false;
        }

        value = TryParseNumeric(selectedValue, out var number)
            ? number
            : string.IsNullOrWhiteSpace(selectedValue)
                ? null
                : selectedValue;

        return true;
    }

    private static bool
        AreDropdownValuesEqual(string? listValue, string selectedValue, string normalizedSelectedValue) {
        var normalizedListValue = NormalizeDropdownText(listValue);

        if (string.Equals(normalizedListValue, normalizedSelectedValue, StringComparison.OrdinalIgnoreCase))
            return true;

        if (string.Equals(
                normalizedListValue.Replace(" ", string.Empty),
                normalizedSelectedValue.Replace(" ", string.Empty),
                StringComparison.OrdinalIgnoreCase))
            return true;

        return !string.IsNullOrWhiteSpace(listValue) &&
               TryParseNumeric(listValue, out var listNumber) &&
               TryParseNumeric(selectedValue, out var selectedNumber) &&
               Math.Abs(listNumber - selectedNumber) < 0.0000001;
    }

    private static string NormalizeDropdownText(string? value) {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var text = value.Replace('\u00A0', ' ');
        var parts = text.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        return string.Join(" ", parts).Trim();
    }

    private static bool TryParseNumeric(string text, out double result) {
        if (string.IsNullOrWhiteSpace(text)) {
            result = 0;
            return false;
        }

        var normalizedValue = text.Trim().Replace(',', '.');

        return double.TryParse(
            normalizedValue,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out result);
    }

    public void RunMacro(string macroName) {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        if (string.IsNullOrWhiteSpace(macroName))
            throw new InvalidOperationException("Nie podano nazwy makra.");

        var macroToRun = macroName.Trim();

        if (!macroToRun.Contains('!')) {
            var workbookName = Convert.ToString(_workbook.Name) ?? string.Empty;
            macroToRun = $"'{workbookName.Replace("'", "''")}'!{macroToRun}";
        }

        _application.Run(macroToRun);
        _application.Calculate();
    }

    public void RunMacroButton(MacroButtonConfig config) {
        if (config == null)
            throw new InvalidOperationException("Nie podano konfiguracji przycisku makra.");

        if (config.IsActiveXCommandButton && !string.IsNullOrWhiteSpace(config.OleObjectName)) {
            RunActiveXCommandButton(config.OleObjectName);
            Recalculate();
            return;
        }

        RunMacro(config.MacroName);
    }

    public void Recalculate() {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        _application.Calculate();
    }

    public void CloseWorkbook() {
        CloseWorkbookInternal();
    }

    private void EnsureApplicationCreated() {
        if (_application != null) return;

        var applicationType = Type.GetTypeFromProgID("Excel.Application");

        if (applicationType == null)
            throw new InvalidOperationException("Nie znaleziono zainstalowanego programu Microsoft Excel.");

        _application = Activator.CreateInstance(applicationType);

        if (_application == null)
            throw new InvalidOperationException("Nie udało się uruchomić programu Microsoft Excel.");

        _application.Visible = false;
        _application.DisplayAlerts = false;
        _application.ScreenUpdating = false;
        _application.EnableEvents = false;
        _application.Interactive = false;

        try {
            _application.Calculation = ExcelCalculationManual;
        }
        catch {
        }

        try {
            _application.AskToUpdateLinks = false;
        }
        catch {
        }

        try {
            _application.DisplayFormulaBar = false;
        }
        catch {
        }

        try {
            _application.DisplayStatusBar = false;
        }
        catch {
        }
    }

    private void OpenWorkbookInternal(string workbookPath) {
        if (_application == null)
            throw new InvalidOperationException("Program Excel nie został uruchomiony.");

        if (!File.Exists(workbookPath))
            throw new FileNotFoundException("Nie znaleziono kopii roboczej kalkulatora.", workbookPath);

        dynamic? workbooks = null;

        try {
            workbooks = _application.Workbooks;

            _workbook = workbooks.Open(
                workbookPath,
                0,
                false,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                true,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                false,
                Type.Missing,
                false);
        }
        finally {
            ReleaseComObject(workbooks);
        }
    }

    private void RunActiveXCommandButton(string oleObjectName) {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        dynamic? worksheet = null;
        dynamic? objects = null;
        dynamic? embeddedObject = null;
        dynamic? control = null;

        try {
            worksheet = GetFirstWorksheet();
            objects = worksheet.OLEObjects();
            embeddedObject = objects.Item(oleObjectName);
            control = embeddedObject.Object;

            try {
                var codeName = Convert.ToString(worksheet.CodeName) ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(codeName)) {
                    var clickMacroName = $"{codeName}.{oleObjectName}_Click";
                    _application.Run(clickMacroName);
                    return;
                }
            }
            catch {
            }

            try {
                control.Value = true;
            }
            catch {
                try {
                    control.Value = false;
                    control.Value = true;
                }
                catch {
                    throw new InvalidOperationException(
                        $"Nie udało się uruchomić przycisku ActiveX '{oleObjectName}'.");
                }
            }
        }
        finally {
            ReleaseComObject(control);
            ReleaseComObject(embeddedObject);
            ReleaseComObject(objects);
            ReleaseComObject(worksheet);
        }
    }

    private void CloseWorkbookInternal() {
        if (_workbook == null) return;

        try {
            _workbook.Close(false);
        }
        catch {
        }
        finally {
            ReleaseComObject(_workbook);
            _workbook = null;
        }
    }

    private static void ReleaseComObject(object? comObject) {
    }
}