using System.Text.Json;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Models;
using DBHWeeklyReport.Core.Services;

namespace DBHWeeklyReport.App;

public sealed class DesktopController(
    ISettingsStore settingsStore,
    IWeeklyReportService weeklyReportService,
    IExcelReportWriter excelReportWriter,
    IAutoStartRegistrar autoStartRegistrar,
    IScheduledTaskRegistrar scheduledTaskRegistrar,
    ILogService logService)
{
    public AppSettings Settings { get; private set; } = new();

    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        Settings = await settingsStore.LoadAsync(cancellationToken);
        if (string.IsNullOrWhiteSpace(Settings.LogDirectory))
        {
            Settings.LogDirectory = Infrastructure.AppPaths.DefaultLogDirectory;
        }
    }

    public ValidationResult Validate() => SettingsValidator.Validate(Settings);

    public async Task SaveAsync(CancellationToken cancellationToken = default)
    {
        await settingsStore.SaveAsync(Settings, cancellationToken);
    }

    public async Task<WeeklyReportPreview> GeneratePreviewAsync(DateOnly? referenceDate = null, CancellationToken cancellationToken = default)
    {
        var preview = await weeklyReportService.GeneratePreviewAsync(Settings, referenceDate, cancellationToken);
        await logService.WriteAsync(
            "preview",
            JsonSerializer.Serialize(preview, new JsonSerializerOptions { WriteIndented = true }),
            cancellationToken);
        return preview;
    }

    public async Task<string> WritePreviewAsync(WeeklyReportPreview preview, CancellationToken cancellationToken = default)
    {
        var path = await excelReportWriter.WriteAsync(Settings.ExcelPath, preview, cancellationToken);
        Settings.LastSuccessfulRunAt = DateTimeOffset.Now;
        await SaveAsync(cancellationToken);
        await logService.WriteAsync(
            "write",
            JsonSerializer.Serialize(preview, new JsonSerializerOptions { WriteIndented = true }),
            cancellationToken);
        return path;
    }

    public string ExecutablePath => Environment.ProcessPath ?? string.Empty;

    public void ApplyAutoStart()
    {
        if (!string.IsNullOrWhiteSpace(ExecutablePath))
        {
            autoStartRegistrar.SetEnabled(ExecutablePath, Settings.AutoStartWithWindows);
        }
    }

    public void ApplyScheduledTask()
    {
        if (!string.IsNullOrWhiteSpace(ExecutablePath))
        {
            scheduledTaskRegistrar.SetEnabled(ExecutablePath, Settings, Settings.UseWindowsTaskScheduler);
        }
    }

    public DateTimeOffset GetNextRun(DateTimeOffset from)
    {
        var target = GetScheduledOccurrenceForWeek(from);
        if (from <= target)
        {
            return target;
        }

        return target.AddDays(7);
    }

    public bool ShouldTrigger(DateTimeOffset now)
    {
        var target = GetScheduledOccurrenceForWeek(now);
        if (now < target)
        {
            return false;
        }

        var weekKey = CreateWeekKey(DateOnly.FromDateTime(target.Date));
        return !string.Equals(Settings.LastTriggeredWeekKey, weekKey, StringComparison.Ordinal);
    }

    public async Task MarkTriggeredAsync(DateOnly referenceDate, CancellationToken cancellationToken = default)
    {
        Settings.LastTriggeredWeekKey = CreateWeekKey(referenceDate);
        Settings.LastPreviewedAt = DateTimeOffset.Now;
        await SaveAsync(cancellationToken);
    }

    public static string CreateWeekKey(DateOnly date)
    {
        var dateTime = date.ToDateTime(TimeOnly.MinValue);
        var isoWeek = System.Globalization.ISOWeek.GetWeekOfYear(dateTime);
        var isoYear = System.Globalization.ISOWeek.GetYear(dateTime);
        return $"{isoYear}-W{isoWeek:00}";
    }

    private DateTimeOffset GetScheduledOccurrenceForWeek(DateTimeOffset now)
    {
        var monday = now.Date.AddDays(-(((int)now.DayOfWeek + 6) % 7));
        var cursor = new DateTimeOffset(monday, now.Offset);
        while (cursor.DayOfWeek != Settings.RunDay)
        {
            cursor = cursor.AddDays(1);
        }

        return cursor.Add(Settings.RunTime);
    }
}
