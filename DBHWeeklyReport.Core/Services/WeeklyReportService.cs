using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Core.Services;

public sealed class WeeklyReportService(
    ICommitCollector commitCollector,
    IWeeklyReportComposer composer,
    ICommitBodyTranslationService commitBodyTranslationService) : IWeeklyReportService
{
    public async Task<WeeklyReportPreview> GeneratePreviewAsync(
        AppSettings settings,
        DateOnly? referenceDate = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(settings.RepositoryPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(settings.ExcelPath);

        var targetDate = referenceDate ?? DateOnly.FromDateTime(DateTime.Today);
        var window = WeeklyReportComposer.ResolveWeekWindow(targetDate);
        var (authorName, authorEmail, commits) = await commitCollector.CollectAsync(
            settings.RepositoryPath,
            window.WeekStart,
            window.WeekEnd,
            string.IsNullOrWhiteSpace(settings.AuthorName) ? null : settings.AuthorName,
            string.IsNullOrWhiteSpace(settings.AuthorEmail) ? null : settings.AuthorEmail,
            cancellationToken);
        var translatedCommits = await commitBodyTranslationService.TranslateAsync(commits, cancellationToken);

        return composer.Compose(
            settings.RepositoryPath,
            settings.ExcelPath,
            authorName,
            authorEmail,
            targetDate,
            translatedCommits);
    }
}
