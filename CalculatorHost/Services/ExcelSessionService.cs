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
        if (comObject == null || !Marshal.IsComObject(comObject)) return;

        try {
            Marshal.FinalReleaseComObject(comObject);
        }
        catch {
        }
    }
}