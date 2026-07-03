using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Services;
using DBHWeeklyReport.Infrastructure;
using WF = System.Windows.Forms;

namespace DBHWeeklyReport.App;

public partial class App : System.Windows.Application
{
    private ServiceProvider? _services;
    private WF.NotifyIcon? _notifyIcon;
    private MainWindow? _mainWindow;
    private ScheduleMonitor? _scheduleMonitor;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        DesktopController? controllerRef = null;
        var services = new ServiceCollection();
        services.AddSingleton<ISettingsStore, JsonSettingsStore>();
        services.AddSingleton<ICommitCollector, GitCommitCollector>();
        services.AddSingleton<ICommitBodyTranslationService, OnlineCommitBodyTranslationService>();
        services.AddSingleton<IWeeklyReportComposer, WeeklyReportComposer>();
        services.AddSingleton<IWeeklyReportService, WeeklyReportService>();
        services.AddSingleton<IExcelReportWriter, PythonExcelReportWriter>();
        services.AddSingleton<IAutoStartRegistrar, RegistryAutoStartRegistrar>();
        services.AddSingleton<IScheduledTaskRegistrar, WindowsScheduledTaskRegistrar>();
        services.AddSingleton<ILogService>(_ => new FileLogService(() =>
            string.IsNullOrWhiteSpace(controllerRef?.Settings.LogDirectory)
                ? AppPaths.DefaultLogDirectory
                : controllerRef.Settings.LogDirectory));
        services.AddSingleton<DesktopController>(provider =>
        {
            controllerRef = new DesktopController(
                provider.GetRequiredService<ISettingsStore>(),
                provider.GetRequiredService<IWeeklyReportService>(),
                provider.GetRequiredService<IExcelReportWriter>(),
                provider.GetRequiredService<IAutoStartRegistrar>(),
                provider.GetRequiredService<IScheduledTaskRegistrar>(),
                provider.GetRequiredService<ILogService>());
            return controllerRef;
        });
        services.AddSingleton<MainWindow>();

        _services = services.BuildServiceProvider();

        var controller = _services.GetRequiredService<DesktopController>();
        await controller.InitializeAsync();

        _mainWindow = _services.GetRequiredService<MainWindow>();
        var startInBackground = e.Args.Any(static arg => string.Equals(arg, "--background", StringComparison.OrdinalIgnoreCase));
        if (!startInBackground)
        {
            _mainWindow.Show();
        }

        ConfigureTray();

        _scheduleMonitor = new ScheduleMonitor(controller);
        _scheduleMonitor.ScheduledPreviewRequested += async referenceDate =>
        {
            if (_mainWindow is null)
            {
                return;
            }

            ShowMainWindow();
            await _mainWindow.ShowPreviewAsync(referenceDate);
        };
        _scheduleMonitor.Start();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _notifyIcon?.Dispose();
        _services?.Dispose();
        base.OnExit(e);
    }

    private void ConfigureTray()
    {
        _notifyIcon = new WF.NotifyIcon
        {
            Text = "Báo cáo tuần DBH",
            Visible = true,
            Icon = System.Drawing.SystemIcons.Application,
        };

        var menu = new WF.ContextMenuStrip();
        menu.Items.Add("Mở", null, (_, _) => ShowMainWindow());
        menu.Items.Add("Chạy ngay", null, async (_, _) =>
        {
            ShowMainWindow();
            if (_mainWindow is not null)
            {
                await _mainWindow.ShowPreviewAsync(DateOnly.FromDateTime(DateTime.Today));
            }
        });
        menu.Items.Add("Thoát", null, (_, _) =>
        {
            _notifyIcon!.Visible = false;
            if (_mainWindow is not null)
            {
                _mainWindow.AllowClose = true;
                _mainWindow.Close();
            }
            Shutdown();
        });

        _notifyIcon.ContextMenuStrip = menu;
        _notifyIcon.DoubleClick += (_, _) => ShowMainWindow();
    }

    private void ShowMainWindow()
    {
        if (_mainWindow is null)
        {
            return;
        }

        _mainWindow.Show();
        _mainWindow.WindowState = WindowState.Normal;
        _mainWindow.Activate();
    }
}
