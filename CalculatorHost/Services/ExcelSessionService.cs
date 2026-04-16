using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalculatorHost.Services;

/// <summary>
///     Manages the lifecycle of a hidden Excel Application instance and its workbook.
///     All methods in this class MUST be called from the ExcelWorker STA thread.
/// </summary>
public class ExcelSessionService : IDisposable {
    private Excel.Application? _application;
    private bool _disposed;
    private Excel.Workbook? _workbook;

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;

        CloseWorkbookInternal();

        if (_application != null)
            try {
                _application.Quit();
            }
            catch {
                // Ignore errors during quit
            }
            finally {
                Marshal.FinalReleaseComObject(_application);
                _application = null;
            }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    public void OpenSession(string workbookPath) {
        EnsureApplicationCreated();
        CloseWorkbookInternal();
        OpenWorkbookInternal(workbookPath);
    }

    private void EnsureApplicationCreated() {
        if (_application != null) return;

        _application = new Excel.Application {
            Visible = false,
            DisplayAlerts = false,
            ScreenUpdating = false,
            EnableEvents = false,
            Interactive = false
        };

        try {
            _application.DisplayFormulaBar = false;
        }
        catch {
            /* Not critical */
        }

        try {
            _application.DisplayStatusBar = false;
        }
        catch {
            /* Not critical */
        }
    }

    private void OpenWorkbookInternal(string workbookPath) {
        if (_application == null) throw new InvalidOperationException("Excel application is not initialized.");

        _workbook = _application.Workbooks.Open(
            workbookPath,
            false,
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

        _application.Calculate();
    }

    public Excel.Worksheet GetFirstWorksheet() {
        if (_workbook == null) throw new InvalidOperationException("No workbook is open.");
        if (_workbook.Worksheets.Count == 0)
            throw new InvalidOperationException("The workbook contains no worksheets.");

        return (Excel.Worksheet)_workbook.Worksheets[1];
    }

    public void WriteCellValue(int row, int column, object? value) {
        if (_application == null || _workbook == null) throw new InvalidOperationException("No active Excel session.");

        Excel.Worksheet? worksheet = null;
        Excel.Range? cell = null;

        try {
            worksheet = GetFirstWorksheet();
            cell = (Excel.Range)worksheet.Cells[row, column];
            cell.Value2 = value;
            _application.Calculate();
        }
        finally {
            ReleaseComObject(ref cell);
            ReleaseComObject(ref worksheet);
        }
    }

    public void RunMacro(string macroName) {
        if (_application == null || _workbook == null) throw new InvalidOperationException("No active Excel session.");

        // Run macro; Excel resolves the name within the open workbook
        _application.Run(macroName);
        _application.Calculate();
    }

    public void Recalculate() {
        _application?.Calculate();
    }

    public void CloseWorkbook() {
        CloseWorkbookInternal();
    }

    private void CloseWorkbookInternal() {
        if (_workbook == null) return;

        try {
            _workbook.Close(false);
        }
        catch {
            // If close fails, force release the COM object anyway
        }
        finally {
            Marshal.FinalReleaseComObject(_workbook);
            _workbook = null;
        }
    }

    private static void ReleaseComObject<T>(ref T? comObject) where T : class {
        if (comObject == null) return;
        Marshal.ReleaseComObject(comObject);
        comObject = null;
    }
}