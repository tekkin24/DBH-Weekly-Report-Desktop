using System.Windows.Threading;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.App;

public sealed class ScheduleMonitor(DesktopController controller)
{
    private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromSeconds(45) };
    private bool _isRunning;

    public event Func<DateOnly, Task>? ScheduledPreviewRequested;

    public void Start()
    {
        _timer.Tick += OnTick;
        _timer.Start();
    }

    private async void OnTick(object? sender, EventArgs e)
    {
        if (_isRunning)
        {
            return;
        }

        if (!controller.Validate().IsValid)
        {
            return;
        }

        var now = DateTimeOffset.Now;
        if (!controller.ShouldTrigger(now))
        {
            return;
        }

        _isRunning = true;
        try
        {
            var referenceDate = DateOnly.FromDateTime(now.DateTime);
            await controller.MarkTriggeredAsync(referenceDate);
            if (ScheduledPreviewRequested is not null)
            {
                await ScheduledPreviewRequested.Invoke(referenceDate);
            }
        }
        finally
        {
            _isRunning = false;
        }
    }
}
