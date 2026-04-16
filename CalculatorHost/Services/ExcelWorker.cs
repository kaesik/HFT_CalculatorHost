using System.Collections.Concurrent;

namespace CalculatorHost.Services;

/// <summary>
///     Maintains a single persistent STA (Single-Threaded Apartment) background thread
///     for all Excel COM interop operations. Excel COM objects must be created and
///     accessed from the same STA thread.
/// </summary>
public class ExcelWorker : IDisposable {
    private readonly BlockingCollection<(Action Action, TaskCompletionSource CompletionSource)> _queue = new();
    private readonly Thread _workerThread;
    private bool _disposed;

    public ExcelWorker() {
        _workerThread = new Thread(ProcessQueue) {
            IsBackground = true,
            Name = "ExcelWorker-STA"
        };
        _workerThread.SetApartmentState(ApartmentState.STA);
        _workerThread.Start();
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _queue.CompleteAdding();
        _workerThread.Join(TimeSpan.FromSeconds(10));
        _queue.Dispose();
    }

    private void ProcessQueue() {
        foreach (var (action, completionSource) in _queue.GetConsumingEnumerable())
            try {
                action();
                completionSource.SetResult();
            }
            catch (Exception exception) {
                completionSource.SetException(exception);
            }
    }

    public Task InvokeAsync(Action action) {
        if (!_disposed) {
            var completionSource = new TaskCompletionSource();
            _queue.Add((action, completionSource));
            return completionSource.Task;
        }

        throw new ObjectDisposedException(nameof(ExcelWorker));
    }

    public async Task<T> InvokeAsync<T>(Func<T> func) {
        T result = default!;
        await InvokeAsync(() => { result = func(); });
        return result;
    }
}