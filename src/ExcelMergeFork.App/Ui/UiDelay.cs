using System.Windows.Threading;

namespace ExcelMergeFork.App.Ui;

/// <summary>
/// Coalesces bursts of UI work, such as filtering on every keystroke.
/// </summary>
public sealed class UiDelay : IDisposable
{
    private readonly DispatcherTimer _timer;
    private Action? _action;

    public UiDelay(int milliseconds = 200)
    {
        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(milliseconds) };
        _timer.Tick += (_, _) =>
        {
            _timer.Stop();
            var action = _action;
            _action = null;
            action?.Invoke();
        };
    }

    public void Schedule(Action action)
    {
        _action = action;
        _timer.Stop();
        _timer.Start();
    }

    public void Dispose()
    {
        _action = null;
        _timer.Stop();
    }
}
