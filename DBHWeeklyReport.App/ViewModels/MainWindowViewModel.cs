using DBHWeeklyReport.App.Support;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.App.ViewModels;

public sealed class MainWindowViewModel : BindableBase
{
    public sealed record RunDayOption(DayOfWeek Value, string Label);

    public static IReadOnlyList<RunDayOption> RunDayOptions { get; } =
    [
        new(DayOfWeek.Monday, "Thứ Hai"),
        new(DayOfWeek.Tuesday, "Thứ Ba"),
        new(DayOfWeek.Wednesday, "Thứ Tư"),
        new(DayOfWeek.Thursday, "Thứ Năm"),
        new(DayOfWeek.Friday, "Thứ Sáu"),
        new(DayOfWeek.Saturday, "Thứ Bảy"),
        new(DayOfWeek.Sunday, "Chủ Nhật"),
    ];

    private string _repositoryPath = string.Empty;
    private string _excelPath = string.Empty;
    private string _authorName = string.Empty;
    private string _authorEmail = string.Empty;
    private string _logDirectory = string.Empty;
    private bool _autoStartWithWindows;
    private bool _useWindowsTaskScheduler;
    private DateTime _selectedReferenceDate = DateTime.Today;
    private string _runTimeText = "16:00";
    private DayOfWeek _selectedRunDay = DayOfWeek.Friday;
    private string _statusText = "Sẵn sàng.";
    private string _nextRunText = string.Empty;
    private string _scheduleText = string.Empty;

    public string RepositoryPath { get => _repositoryPath; set => SetProperty(ref _repositoryPath, value); }
    public string ExcelPath { get => _excelPath; set => SetProperty(ref _excelPath, value); }
    public string AuthorName { get => _authorName; set => SetProperty(ref _authorName, value); }
    public string AuthorEmail { get => _authorEmail; set => SetProperty(ref _authorEmail, value); }
    public string LogDirectory { get => _logDirectory; set => SetProperty(ref _logDirectory, value); }
    public bool AutoStartWithWindows { get => _autoStartWithWindows; set => SetProperty(ref _autoStartWithWindows, value); }
    public bool UseWindowsTaskScheduler { get => _useWindowsTaskScheduler; set => SetProperty(ref _useWindowsTaskScheduler, value); }
    public DateTime SelectedReferenceDate { get => _selectedReferenceDate; set => SetProperty(ref _selectedReferenceDate, value); }
    public string RunTimeText { get => _runTimeText; set => SetProperty(ref _runTimeText, value); }
    public DayOfWeek SelectedRunDay { get => _selectedRunDay; set => SetProperty(ref _selectedRunDay, value); }
    public IReadOnlyList<RunDayOption> AvailableRunDays => RunDayOptions;
    public string StatusText { get => _statusText; set => SetProperty(ref _statusText, value); }
    public string NextRunText { get => _nextRunText; set => SetProperty(ref _nextRunText, value); }
    public string ScheduleText { get => _scheduleText; set => SetProperty(ref _scheduleText, value); }

    public void ApplySettings(AppSettings settings)
    {
        RepositoryPath = settings.RepositoryPath;
        ExcelPath = settings.ExcelPath;
        AuthorName = settings.AuthorName;
        AuthorEmail = settings.AuthorEmail;
        LogDirectory = settings.LogDirectory;
        AutoStartWithWindows = settings.AutoStartWithWindows;
        UseWindowsTaskScheduler = settings.UseWindowsTaskScheduler;
        RunTimeText = settings.RunTime.ToString(@"hh\:mm");
        SelectedRunDay = settings.RunDay;
    }

    public void UpdateSettings(AppSettings settings)
    {
        settings.RepositoryPath = RepositoryPath;
        settings.ExcelPath = ExcelPath;
        settings.AuthorName = AuthorName;
        settings.AuthorEmail = AuthorEmail;
        settings.LogDirectory = LogDirectory;
        settings.AutoStartWithWindows = AutoStartWithWindows;
        settings.UseWindowsTaskScheduler = UseWindowsTaskScheduler;
        settings.RunDay = SelectedRunDay;
        if (TimeSpan.TryParse(RunTimeText, out var parsedTime))
        {
            settings.RunTime = parsedTime;
        }
    }
}
