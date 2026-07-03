using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Core;

public interface ICommitCollector
{
    Task<(string? AuthorName, string? AuthorEmail, IReadOnlyList<CommitInfo> Commits)> CollectAsync(
        string repositoryPath,
        DateOnly startDate,
        DateOnly endDate,
        string? authorName,
        string? authorEmail,
        CancellationToken cancellationToken = default);
}

public interface IWeeklyReportComposer
{
    WeeklyReportPreview Compose(
        string repositoryPath,
        string excelPath,
        string? authorName,
        string? authorEmail,
        DateOnly referenceDate,
        IReadOnlyList<CommitInfo> commits);
}

public interface IWeeklyReportService
{
    Task<WeeklyReportPreview> GeneratePreviewAsync(
        AppSettings settings,
        DateOnly? referenceDate = null,
        CancellationToken cancellationToken = default);
}

public interface ICommitBodyTranslationService
{
    Task<IReadOnlyList<CommitInfo>> TranslateAsync(
        IReadOnlyList<CommitInfo> commits,
        CancellationToken cancellationToken = default);
}

public interface IExcelReportWriter
{
    Task<string> WriteAsync(
        string excelPath,
        WeeklyReportPreview preview,
        CancellationToken cancellationToken = default);
}

public interface ISettingsStore
{
    Task<AppSettings> LoadAsync(CancellationToken cancellationToken = default);

    Task SaveAsync(AppSettings settings, CancellationToken cancellationToken = default);
}

public interface IAutoStartRegistrar
{
    bool IsEnabled(string executablePath);

    void SetEnabled(string executablePath, bool enabled);
}

public interface ILogService
{
    string LogDirectory { get; }

    Task<string> WriteAsync(string category, string content, CancellationToken cancellationToken = default);
}
