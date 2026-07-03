using System.Diagnostics;
using System.Runtime.Versioning;
using System.Text;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Infrastructure;

[SupportedOSPlatform("windows")]
public sealed class WindowsScheduledTaskRegistrar : IScheduledTaskRegistrar
{
    private const string TaskName = "DBHWeeklyReportDesktopWeeklyRun";
    private const string BackgroundArgument = "--background";

    public bool IsEnabled(string executablePath, AppSettings settings)
    {
        var result = RunSchtasks("/Query", "/TN", TaskName);
        return result.ExitCode == 0 && result.StdOut.Contains(TaskName, StringComparison.OrdinalIgnoreCase);
    }

    public void SetEnabled(string executablePath, AppSettings settings, bool enabled)
    {
        if (enabled)
        {
            var dayCode = settings.RunDay switch
            {
                DayOfWeek.Monday => "MON",
                DayOfWeek.Tuesday => "TUE",
                DayOfWeek.Wednesday => "WED",
                DayOfWeek.Thursday => "THU",
                DayOfWeek.Friday => "FRI",
                DayOfWeek.Saturday => "SAT",
                DayOfWeek.Sunday => "SUN",
                _ => "FRI",
            };

            var runTime = settings.RunTime.ToString(@"hh\:mm");
            var command = $"\"{executablePath}\" {BackgroundArgument}";
            var result = RunSchtasks(
                "/Create",
                "/F",
                "/TN",
                TaskName,
                "/SC",
                "WEEKLY",
                "/D",
                dayCode,
                "/ST",
                runTime,
                "/TR",
                command,
                "/IT");

            if (result.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    $"Không thể tạo lịch Windows Task Scheduler.{Environment.NewLine}{result.StdOut}{Environment.NewLine}{result.StdErr}");
            }

            return;
        }

        var deleteResult = RunSchtasks("/Delete", "/F", "/TN", TaskName);
        if (deleteResult.ExitCode != 0 &&
            !deleteResult.StdErr.Contains("cannot find the file", StringComparison.OrdinalIgnoreCase) &&
            !deleteResult.StdOut.Contains("cannot find the file", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Không thể xóa lịch Windows Task Scheduler.{Environment.NewLine}{deleteResult.StdOut}{Environment.NewLine}{deleteResult.StdErr}");
        }
    }

    private static (int ExitCode, string StdOut, string StdErr) RunSchtasks(params string[] arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "schtasks.exe",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };

        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = new Process { StartInfo = startInfo };
        process.Start();
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return (process.ExitCode, stdout, stderr);
    }
}
