namespace DBHWeeklyReport.Core.Models;

public sealed record CommitInfo(
    DateOnly CommitDate,
    string Subject,
    IReadOnlyList<string> BodyLines,
    IReadOnlyList<string> Files);

public sealed class DailyReportEntry
{
    public DateOnly ReportDate { get; init; }

    public string Status { get; init; } = string.Empty;

    public string Task { get; init; } = string.Empty;

    public string Detail { get; init; } = string.Empty;

    public string Memo { get; init; } = string.Empty;

    public int CommitCount { get; init; }
}

public sealed class WeeklyReportPreview
{
    public required string RepositoryPath { get; init; }

    public required string ExcelPath { get; init; }

    public string? AuthorName { get; init; }

    public string? AuthorEmail { get; init; }

    public required DateOnly WeekStart { get; init; }

    public required DateOnly WeekEnd { get; init; }

    public required int CommitCount { get; init; }

    public required IReadOnlyList<DailyReportEntry> Reports { get; init; }

    public IReadOnlyList<string> TouchedSheets { get; init; } = Array.Empty<string>();
}

public sealed class ValidationResult
{
    public bool IsValid => Errors.Count == 0;

    public List<string> Errors { get; } = [];
}

public sealed class WeekWindow
{
    public required DateOnly WeekStart { get; init; }

    public required DateOnly WeekEnd { get; init; }
}
