using System.Windows;
using CalculatorHost.Services;

namespace CalculatorHost;

public partial class App {
    /// <summary>
    ///     Single shared STA worker thread for all Excel COM operations.
    ///     One ExcelWorker per application process is correct — sessions are isolated by ExcelSessionService.
    /// </summary>
    public static ExcelWorker ExcelWorker { get; private set; } = null!;

    protected override void OnStartup(StartupEventArgs e) {
        base.OnStartup(e);
        ThemeService.Initialize();
        ExcelWorker = new ExcelWorker();

        var mainWindow = new MainWindow();
        MainWindow = mainWindow;
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e) {
        ExcelWorker.Dispose();
        base.OnExit(e);
    }
}