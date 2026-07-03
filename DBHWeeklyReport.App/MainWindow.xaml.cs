using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using DBHWeeklyReport.App.ViewModels;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.App;

public partial class MainWindow : Window
{
    private readonly DesktopController _controller;
    private readonly MainWindowViewModel _viewModel;

    public MainWindow(DesktopController controller)
    {
        InitializeComponent();
        _controller = controller;
        _viewModel = new MainWindowViewModel();
        DataContext = _viewModel;
        Loaded += OnLoaded;
        Closing += OnClosing;
    }

    public bool AllowClose { get; set; }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        _viewModel.ApplySettings(_controller.Settings);
        RefreshSchedule();
    }

    private void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (AllowClose)
        {
            return;
        }

        e.Cancel = true;
        Hide();
    }

    public async Task<bool> ShowPreviewAsync(DateOnly referenceDate)
    {
        try
        {
            _viewModel.StatusText = "Đang tạo bản xem trước...";
            var preview = await _controller.GeneratePreviewAsync(referenceDate);
            var dialog = new PreviewWindow(new PreviewWindowViewModel(preview));
            var result = dialog.ShowDialog();
            if (result == true && dialog.Confirmed)
            {
                _viewModel.StatusText = "Đang ghi vào Excel...";
                var path = await _controller.WritePreviewAsync(preview);
                _viewModel.StatusText = $"Đã ghi vào Excel: {path}";
                RefreshSchedule();
                return true;
            }

            _viewModel.StatusText = "Đã hủy xem trước.";
            return false;
        }
        catch (Exception ex)
        {
            _viewModel.StatusText = ex.Message;
            System.Windows.MessageBox.Show(this, ex.ToString(), "Lỗi báo cáo tuần", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    private async void OnSaveClick(object sender, RoutedEventArgs e)
    {
        try
        {
            SyncSettingsFromView();
            var validation = _controller.Validate();
            if (!validation.IsValid)
            {
                System.Windows.MessageBox.Show(this, string.Join(Environment.NewLine, validation.Errors), "Kiểm tra dữ liệu", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await _controller.SaveAsync();
            _controller.ApplyAutoStart();
            _controller.ApplyScheduledTask();
            _viewModel.StatusText = "Đã lưu cài đặt.";
            RefreshSchedule();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.ToString(), "Lỗi lưu cài đặt", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void OnRunNowClick(object sender, RoutedEventArgs e)
    {
        SyncSettingsFromView();
        await ShowPreviewAsync(DateOnly.FromDateTime(DateTime.Today));
    }

    private async void OnPreviewSelectedWeekClick(object sender, RoutedEventArgs e)
    {
        SyncSettingsFromView();
        await ShowPreviewAsync(DateOnly.FromDateTime(_viewModel.SelectedReferenceDate));
    }

    private void OnBrowseRepositoryClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            _viewModel.RepositoryPath = dialog.SelectedPath;
        }
    }

    private void OnBrowseLogDirectoryClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            _viewModel.LogDirectory = dialog.SelectedPath;
        }
    }

    private void OnBrowseExcelClick(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Tệp Excel (*.xlsx)|*.xlsx",
            CheckFileExists = true,
        };

        if (dialog.ShowDialog(this) == true)
        {
            _viewModel.ExcelPath = dialog.FileName;
        }
    }

    private void OnOpenLogsClick(object sender, RoutedEventArgs e)
    {
        SyncSettingsFromView();
        Directory.CreateDirectory(_controller.Settings.LogDirectory);
        Process.Start(new ProcessStartInfo
        {
            FileName = _controller.Settings.LogDirectory,
            UseShellExecute = true,
        });
    }

    private void SyncSettingsFromView()
    {
        _viewModel.UpdateSettings(_controller.Settings);
    }

    private void RefreshSchedule()
    {
        var nextRun = _controller.GetNextRun(DateTimeOffset.Now);
        _viewModel.NextRunText = $"{GetVietnameseDayName(nextRun.DayOfWeek)}, {nextRun:dd/MM/yyyy HH:mm}";
        _viewModel.ScheduleText = $"Mỗi {GetVietnameseDayName(_controller.Settings.RunDay)} lúc {_controller.Settings.RunTime:hh\\:mm}";
    }

    private static string GetVietnameseDayName(DayOfWeek dayOfWeek) => dayOfWeek switch
    {
        DayOfWeek.Monday => "Thứ Hai",
        DayOfWeek.Tuesday => "Thứ Ba",
        DayOfWeek.Wednesday => "Thứ Tư",
        DayOfWeek.Thursday => "Thứ Năm",
        DayOfWeek.Friday => "Thứ Sáu",
        DayOfWeek.Saturday => "Thứ Bảy",
        DayOfWeek.Sunday => "Chủ Nhật",
        _ => dayOfWeek.ToString(),
    };
}
