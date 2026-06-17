namespace DBHWeeklyReport.Core.Models;

public sealed class AppSettings
{
    public string RepositoryPath { get; set; } = string.Empty;

    public string ExcelPath { get; set; } = string.Empty;

    public string AuthorName { get; set; } = string.Empty;

    public string AuthorEmail { get; set; } = string.Empty;

    public DayOfWeek RunDay { get; set; } = DayOfWeek.Friday;

    public TimeSpan RunTime { get; set; } = new(16, 0, 0);

    public bool AutoStartWithWindows { get; set; } = true;

    public string LogDirectory { get; set; } = string.Empty;

    public string LastTriggeredWeekKey { get; set; } = string.Empty;

    public DateTimeOffset? LastSuccessfulRunAt { get; set; }

    public DateTimeOffset? LastPreviewedAt { get; set; }
}
