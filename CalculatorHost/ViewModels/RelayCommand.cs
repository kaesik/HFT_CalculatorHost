using System.Windows.Input;

namespace CalculatorHost.ViewModels;

public class RelayCommand : ICommand {
    private readonly Func<object?, bool>? _canExecute;
    private readonly Action<object?> _execute;

    private RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null) {
        _execute = execute;
        _canExecute = canExecute;
    }

    public RelayCommand(Action execute, Func<bool>? canExecute = null)
        : this(_ => execute(), canExecute == null ? null : _ => canExecute()) {
    }

    public event EventHandler? CanExecuteChanged {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter) {
        return _canExecute?.Invoke(parameter) ?? true;
    }

    public void Execute(object? parameter) {
        _execute(parameter);
    }
}

public class AsyncRelayCommand : ICommand {
    private readonly Func<object?, bool>? _canExecute;
    private readonly Func<object?, Task> _execute;
    private bool _isExecuting;

    public AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null)
        : this(_ => execute(), canExecute == null ? null : _ => canExecute()) {
    }

    private AsyncRelayCommand(Func<object?, Task> execute, Func<object?, bool>? canExecute = null) {
        _execute = execute;
        _canExecute = canExecute;
    }

    public event EventHandler? CanExecuteChanged {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter) {
        return !_isExecuting && (_canExecute?.Invoke(parameter) ?? true);
    }

    public async void Execute(object? parameter) {
        try {
            _isExecuting = true;
            CommandManager.InvalidateRequerySuggested();
            try {
                await _execute(parameter);
            }
            finally {
                _isExecuting = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }
        catch {
            // ignored
        }
    }
}

public class AsyncRelayCommand<T>(Func<T?, Task> execute, Func<T?, bool>? canExecute = null) : ICommand {
    private bool _isExecuting;

    public event EventHandler? CanExecuteChanged {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter) {
        return !_isExecuting && (canExecute?.Invoke(parameter is T t ? t : default) ?? true);
    }

    public async void Execute(object? parameter) {
        try {
            _isExecuting = true;
            CommandManager.InvalidateRequerySuggested();
            try {
                await execute(parameter is T t ? t : default);
            }
            finally {
                _isExecuting = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }
        catch {
            // ignored
        }
    }
}