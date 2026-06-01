using System.IO;
using System.Runtime.InteropServices;
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

    public void WriteCellValue(int row, int column, object? value) {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        dynamic? worksheet = null;
        dynamic? cells = null;
        dynamic? cell = null;

        try {
            worksheet = GetFirstWorksheet();
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

    public void RunMacro(string macroName) {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        if (string.IsNullOrWhiteSpace(macroName))
            throw new InvalidOperationException("Nie podano nazwy makra.");

        var normalizedMacroName = macroName.Trim();
        var macroToRun = normalizedMacroName.Contains('!')
            ? normalizedMacroName
            : $"'{(Convert.ToString(_workbook.Name) ?? string.Empty).Replace("'", "''")}'!{normalizedMacroName}";

        _application.Run(macroToRun);
        _application.Calculate();
    }

    public void RunMacroButton(MacroButtonConfig config) {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        if (config.ActionType == MacroButtonActionType.ActiveXClick) {
            ClickActiveXButton(config);
            _application.Calculate();
            return;
        }

        RunMacro(config.MacroName);
    }

    private void ClickActiveXButton(MacroButtonConfig config) {
        if (_application == null || _workbook == null)
            throw new InvalidOperationException("Brak aktywnej sesji programu Excel.");

        if (string.IsNullOrWhiteSpace(config.OleObjectName))
            throw new InvalidOperationException("Nie podano nazwy przycisku ActiveX.");

        dynamic? worksheets = null;
        dynamic? worksheet = null;
        dynamic? oleObjects = null;
        dynamic? oleObject = null;
        dynamic? control = null;
        object? previousEnableEvents = null;

        try {
            worksheets = _workbook.Worksheets;
            worksheet = string.IsNullOrWhiteSpace(config.SheetName)
                ? worksheets[1]
                : worksheets[config.SheetName];

            oleObjects = worksheet.OLEObjects();
            oleObject = oleObjects.Item(config.OleObjectName);
            control = oleObject.Object;

            try {
                previousEnableEvents = _application.EnableEvents;
                _application.EnableEvents = true;
            }
            catch {
            }

            try {
                control.Value = true;
                control.Value = false;
            }
            catch {
                control.Value = true;
            }
        }
        finally {
            if (previousEnableEvents != null)
                try {
                    _application.EnableEvents = previousEnableEvents;
                }
                catch {
                }

            ReleaseComObject(control);
            ReleaseComObject(oleObject);
            ReleaseComObject(oleObjects);
            ReleaseComObject(worksheet);
            ReleaseComObject(worksheets);
        }
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
        if (comObject == null || !Marshal.IsComObject(comObject)) return;

        try {
            Marshal.FinalReleaseComObject(comObject);
        }
        catch {
        }
    }
}