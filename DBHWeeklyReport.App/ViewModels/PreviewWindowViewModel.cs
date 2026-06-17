using System.Collections.ObjectModel;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.App.ViewModels;

public sealed class PreviewWindowViewModel
{
    public PreviewWindowViewModel(WeeklyReportPreview preview)
    {
        Preview = preview;
        Items = new ObservableCollection<DailyReportEntry>(preview.Reports);
        Summary = $"{preview.WeekStart:yyyy-MM-dd} - {preview.WeekEnd:yyyy-MM-dd} | Commits: {preview.CommitCount} | Days: {preview.Reports.Count}";
    }

    public WeeklyReportPreview Preview { get; }

    public ObservableCollection<DailyReportEntry> Items { get; }

    public string Summary { get; }
}
