using System.Text.Json;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Infrastructure;

public sealed class JsonSettingsStore : ISettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    public async Task<AppSettings> LoadAsync(CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(AppPaths.AppDataRoot);

        if (!File.Exists(AppPaths.SettingsPath))
        {
            var defaults = CreateDefaultSettings();
            EnvironmentProbe.ApplyDefaults(defaults);
            return defaults;
        }

        await using var stream = File.OpenRead(AppPaths.SettingsPath);
        var settings = await JsonSerializer.DeserializeAsync<AppSettings>(stream, JsonOptions, cancellationToken);
        settings ??= CreateDefaultSettings();
        EnvironmentProbe.ApplyDefaults(settings);
        return settings;
    }

    public async Task SaveAsync(AppSettings settings, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(AppPaths.AppDataRoot);
        await using var stream = File.Create(AppPaths.SettingsPath);
        await JsonSerializer.SerializeAsync(stream, settings, JsonOptions, cancellationToken);
    }

    private static AppSettings CreateDefaultSettings() => new()
    {
        RunDay = DayOfWeek.Friday,
        RunTime = new TimeSpan(16, 0, 0),
        AutoStartWithWindows = true,
        UseWindowsTaskScheduler = true,
        LogDirectory = AppPaths.DefaultLogDirectory,
    };
}
