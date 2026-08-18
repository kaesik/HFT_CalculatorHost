using System.Threading;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

public sealed record CalculatorCacheUpdateProgress(
    int CurrentIndex,
    int TotalCount,
    string CalculatorName,
    string Stage,
    string Detail,
    double OverallPercentage,
    double StagePercentage,
    int UpdatedCount,
    int SkippedCount,
    int FailedCount);

public sealed record CalculatorCacheUpdateResult(
    int UpdatedCount,
    int SkippedCount,
    int FailedCount,
    bool WasCancelled);

public sealed class CalculatorCacheUpdateService {
    private readonly SheetLayoutCacheService _cacheService;
    private readonly ExcelWorker _worker;

    public CalculatorCacheUpdateService(
        SheetLayoutCacheService cacheService,
        ExcelWorker worker) {
        _cacheService = cacheService;
        _worker = worker;
    }

    public async Task<CalculatorCacheUpdateResult> UpdateAllAsync(
        IReadOnlyList<CalculatorInfo> calculators,
        IProgress<CalculatorCacheUpdateProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        if (calculators.Count == 0)
            return new CalculatorCacheUpdateResult(0, 0, 0, false);

        var session = new ExcelSessionService();
        var reader = new SheetReaderService();
        var workingCopy = new WorkingCopyService();

        var updatedCount = 0;
        var skippedCount = 0;
        var failedCount = 0;
        var wasCancelled = false;

        try {
            for (var index = 0; index < calculators.Count; index++) {
                if (cancellationToken.IsCancellationRequested) {
                    wasCancelled = true;
                    break;
                }

                var calculator = calculators[index];
                var currentIndex = index + 1;

                if (_cacheService.IsCacheCurrent(calculator)) {
                    skippedCount++;

                    Report(
                        progress,
                        currentIndex,
                        calculators.Count,
                        calculator,
                        "Cache aktualny",
                        "Pominięto — cache nie wymaga przebudowy.",
                        GetOverallPercentage(index, 100.0, calculators.Count),
                        100.0,
                        updatedCount,
                        skippedCount,
                        failedCount);

                    continue;
                }

                try {
                    Report(
                        progress,
                        currentIndex,
                        calculators.Count,
                        calculator,
                        "Przygotowanie pliku",
                        "Tworzenie kopii roboczej…",
                        GetOverallPercentage(index, 2.0, calculators.Count),
                        0.0,
                        updatedCount,
                        skippedCount,
                        failedCount);

                    var workingPath = workingCopy.CreateWorkingCopy(calculator.FilePath);

                    Report(
                        progress,
                        currentIndex,
                        calculators.Count,
                        calculator,
                        "Uruchamianie Excela",
                        "Otwieranie skoroszytu…",
                        GetOverallPercentage(index, 8.0, calculators.Count),
                        0.0,
                        updatedCount,
                        skippedCount,
                        failedCount);

                    await _worker.InvokeAsync(() => session.OpenSession(workingPath));

                    Report(
                        progress,
                        currentIndex,
                        calculators.Count,
                        calculator,
                        "Uruchamianie Excela",
                        "Skoroszyt otwarty.",
                        GetOverallPercentage(index, 15.0, calculators.Count),
                        100.0,
                        updatedCount,
                        skippedCount,
                        failedCount);

                    var sheetProgress = new Progress<SheetReadProgress>(sheetProgressValue => {
                        var fileProgress = 15.0 +
                                           Math.Clamp(sheetProgressValue.OverallPercentage, 0.0, 100.0) * 0.75;

                        Report(
                            progress,
                            currentIndex,
                            calculators.Count,
                            calculator,
                            sheetProgressValue.Stage,
                            sheetProgressValue.Detail,
                            GetOverallPercentage(index, fileProgress, calculators.Count),
                            sheetProgressValue.StagePercentage,
                            updatedCount,
                            skippedCount,
                            failedCount);
                    });

                    var model = await _worker.InvokeAsync(() =>
                        reader.ReadFirstSheet(session, sheetProgress));

                    if (cancellationToken.IsCancellationRequested) {
                        wasCancelled = true;
                        break;
                    }

                    Report(
                        progress,
                        currentIndex,
                        calculators.Count,
                        calculator,
                        "Zapisywanie cache",
                        "Zapisywanie nowego układu…",
                        GetOverallPercentage(index, 92.0, calculators.Count),
                        0.0,
                        updatedCount,
                        skippedCount,
                        failedCount);

                    if (!_cacheService.TrySave(calculator, model))
                        throw new InvalidOperationException("Nie udało się zapisać cache kalkulatora.");

                    updatedCount++;

                    Report(
                        progress,
                        currentIndex,
                        calculators.Count,
                        calculator,
                        "Gotowe",
                        "Cache kalkulatora został zaktualizowany.",
                        GetOverallPercentage(index, 100.0, calculators.Count),
                        100.0,
                        updatedCount,
                        skippedCount,
                        failedCount);
                }
                catch (Exception exception) {
                    failedCount++;

                    Report(
                        progress,
                        currentIndex,
                        calculators.Count,
                        calculator,
                        "Błąd",
                        GetExceptionMessage(exception),
                        GetOverallPercentage(index, 100.0, calculators.Count),
                        100.0,
                        updatedCount,
                        skippedCount,
                        failedCount);
                }
                finally {
                    try {
                        await _worker.InvokeAsync(session.CloseWorkbook);
                    }
                    catch {
                        // Best-effort cleanup.
                    }

                    workingCopy.CleanCurrentSession();
                }
            }
        }
        finally {
            workingCopy.CleanCurrentSession();

            try {
                await _worker.InvokeAsync(session.Dispose);
            }
            catch {
                // Best-effort cleanup.
            }

            workingCopy.Dispose();
        }

        return new CalculatorCacheUpdateResult(
            updatedCount,
            skippedCount,
            failedCount,
            wasCancelled || cancellationToken.IsCancellationRequested);
    }

    private static double GetOverallPercentage(
        int zeroBasedCalculatorIndex,
        double currentCalculatorPercentage,
        int calculatorCount) {
        if (calculatorCount <= 0)
            return 100.0;

        var completedPart = zeroBasedCalculatorIndex;
        var currentPart = Math.Clamp(currentCalculatorPercentage, 0.0, 100.0) / 100.0;

        return Math.Clamp(
            (completedPart + currentPart) * 100.0 / calculatorCount,
            0.0,
            100.0);
    }

    private static void Report(
        IProgress<CalculatorCacheUpdateProgress>? progress,
        int currentIndex,
        int totalCount,
        CalculatorInfo calculator,
        string stage,
        string detail,
        double overallPercentage,
        double stagePercentage,
        int updatedCount,
        int skippedCount,
        int failedCount) {
        progress?.Report(new CalculatorCacheUpdateProgress(
            currentIndex,
            totalCount,
            calculator.DisplayName,
            stage,
            detail,
            Math.Clamp(overallPercentage, 0.0, 100.0),
            Math.Clamp(stagePercentage, 0.0, 100.0),
            updatedCount,
            skippedCount,
            failedCount));
    }

    private static string GetExceptionMessage(Exception exception) {
        var messages = new List<string>();
        Exception? currentException = exception;

        while (currentException != null) {
            if (!string.IsNullOrWhiteSpace(currentException.Message))
                messages.Add(currentException.Message);

            currentException = currentException.InnerException;
        }

        return messages.Count == 0
            ? "Nieznany błąd."
            : string.Join(" | ", messages);
    }
}
