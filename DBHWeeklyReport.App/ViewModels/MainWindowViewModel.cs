using DBHWeeklyReport.App.Support;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.App.ViewModels;

public sealed class MainWindowViewModel : BindableBase
{
    private string _repositoryPath = string.Empty;
    private string _excelPath = string.Empty;
    private string _authorName = string.Empty;
    private string _authorEmail = string.Empty;
    private string _logDirectory = string.Empty;
    private bool _autoStartWithWindows;
    private DateTime _selectedReferenceDate = DateTime.Today;
    private string _runTimeText = "16:00";
    private string _statusText = "Ready.";
    private string _nextRunText = string.Empty;

    public string RepositoryPath { get => _repositoryPath; set => SetProperty(ref _repositoryPath, value); }
    public string ExcelPath { get => _excelPath; set => SetProperty(ref _excelPath, value); }
    public string AuthorName { get => _authorName; set => SetProperty(ref _authorName, value); }
    public string AuthorEmail { get => _authorEmail; set => SetProperty(ref _authorEmail, value); }
    public string LogDirectory { get => _logDirectory; set => SetProperty(ref _logDirectory, value); }
    public bool AutoStartWithWindows { get => _autoStartWithWindows; set => SetProperty(ref _autoStartWithWindows, value); }
    public DateTime SelectedReferenceDate { get => _selectedReferenceDate; set => SetProperty(ref _selectedReferenceDate, value); }
    public string RunTimeText { get => _runTimeText; set => SetProperty(ref _runTimeText, value); }
    public string StatusText { get => _statusText; set => SetProperty(ref _statusText, value); }
    public string NextRunText { get => _nextRunText; set => SetProperty(ref _nextRunText, value); }

    public void ApplySettings(AppSettings settings)
    {
        RepositoryPath = settings.RepositoryPath;
        ExcelPath = settings.ExcelPath;
        AuthorName = settings.AuthorName;
        AuthorEmail = settings.AuthorEmail;
        LogDirectory = settings.LogDirectory;
        AutoStartWithWindows = settings.AutoStartWithWindows;
        RunTimeText = settings.RunTime.ToString(@"hh\:mm");
    }

    public void UpdateSettings(AppSettings settings)
    {
        settings.RepositoryPath = RepositoryPath;
        settings.ExcelPath = ExcelPath;
        settings.AuthorName = AuthorName;
        settings.AuthorEmail = AuthorEmail;
        settings.LogDirectory = LogDirectory;
        settings.AutoStartWithWindows = AutoStartWithWindows;
        if (TimeSpan.TryParse(RunTimeText, out var parsedTime))
        {
            settings.RunTime = parsedTime;
        }
    }
}
