using System.Diagnostics;
using System.Text;
using System.Text.Json;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Infrastructure;

public sealed class PythonExcelReportWriter : IExcelReportWriter
{
    private readonly string _scriptPath;

    public PythonExcelReportWriter()
    {
        var bundledScript = Path.Combine(AppContext.BaseDirectory, "weekly_report", "fill_weekly_report.py");
        if (File.Exists(bundledScript))
        {
            _scriptPath = bundledScript;
            return;
        }

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

        var previewJsonPath = Path.Combine(Path.GetTempPath(), $"dbh-weekly-preview-{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(
            previewJsonPath,
            JsonSerializer.Serialize(
                preview.Reports.Select(static report => new SerializableDailyReportEntry
                {
                    ReportDate = report.ReportDate.ToString("yyyy-MM-dd"),
                    Status = report.Status,
                    Task = report.Task,
                    Detail = report.Detail,
                    Memo = report.Memo,
                    CommitCount = report.CommitCount,
                })),
            new UTF8Encoding(false),
            cancellationToken);

        try
        {
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
            startInfo.ArgumentList.Add("--preview-json");
            startInfo.ArgumentList.Add(previewJsonPath);

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
        finally
        {
            try
            {
                if (File.Exists(previewJsonPath))
                {
                    File.Delete(previewJsonPath);
                }
            }
            catch
            {
            }
        }
    }

    private sealed class SerializableDailyReportEntry
    {
        public required string ReportDate { get; init; }

        public required string Status { get; init; }

        public required string Task { get; init; }

        public required string Detail { get; init; }

        public required string Memo { get; init; }

        public required int CommitCount { get; init; }
    }
}
