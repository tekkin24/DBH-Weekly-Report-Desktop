using System.Diagnostics;
using System.Text;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Infrastructure;

public sealed class PythonExcelReportWriter : IExcelReportWriter
{
    private readonly string _scriptPath;

    public PythonExcelReportWriter()
    {
        _scriptPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Documents",
            "DBH Weekly Report Automation",
            "weekly_report",
            "fill_weekly_report.py");
    }

    public async Task<string> WriteAsync(string excelPath, WeeklyReportPreview preview, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(_scriptPath))
        {
            throw new FileNotFoundException("Weekly report Python script was not found.", _scriptPath);
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = "python.exe",
            WorkingDirectory = preview.RepositoryPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };

        startInfo.Environment["PYTHONIOENCODING"] = "utf-8";
        startInfo.ArgumentList.Add(_scriptPath);
        startInfo.ArgumentList.Add("--repo-path");
        startInfo.ArgumentList.Add(preview.RepositoryPath);
        startInfo.ArgumentList.Add("--excel-path");
        startInfo.ArgumentList.Add(excelPath);
        startInfo.ArgumentList.Add("--reference-date");
        startInfo.ArgumentList.Add(preview.WeekEnd.ToString("yyyy-MM-dd"));

        if (!string.IsNullOrWhiteSpace(preview.AuthorName))
        {
            startInfo.ArgumentList.Add("--author");
            startInfo.ArgumentList.Add(preview.AuthorName);
        }

        if (!string.IsNullOrWhiteSpace(preview.AuthorEmail))
        {
            startInfo.ArgumentList.Add("--author-email");
            startInfo.ArgumentList.Add(preview.AuthorEmail);
        }

        using var process = new Process { StartInfo = startInfo };
        process.Start();
        var stdout = await process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderr = await process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"Python Excel writer failed with exit code {process.ExitCode}.{Environment.NewLine}{stdout}{Environment.NewLine}{stderr}");
        }

        return excelPath;
    }
}
