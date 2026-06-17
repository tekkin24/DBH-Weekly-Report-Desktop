using System.Diagnostics;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Infrastructure;

public static class EnvironmentProbe
{
    public static void ApplyDefaults(AppSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.LogDirectory))
        {
            settings.LogDirectory = AppPaths.DefaultLogDirectory;
        }

        var oneDriveRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "OneDrive");

        if (Directory.Exists(oneDriveRoot))
        {
            if (string.IsNullOrWhiteSpace(settings.RepositoryPath))
            {
                settings.RepositoryPath = FindRepositoryRoot(oneDriveRoot) ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(settings.ExcelPath))
            {
                settings.ExcelPath = FindExcelPath(oneDriveRoot) ?? string.Empty;
            }
        }

        if (!string.IsNullOrWhiteSpace(settings.RepositoryPath))
        {
            if (string.IsNullOrWhiteSpace(settings.AuthorName))
            {
                settings.AuthorName = TryReadGitConfig(settings.RepositoryPath, "user.name") ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(settings.AuthorEmail))
            {
                settings.AuthorEmail = TryReadGitConfig(settings.RepositoryPath, "user.email") ?? string.Empty;
            }
        }
    }

    private static string? FindRepositoryRoot(string oneDriveRoot)
    {
        return Directory.EnumerateFiles(oneDriveRoot, "DBH.slnx", SearchOption.AllDirectories)
            .Select(Path.GetDirectoryName)
            .FirstOrDefault(static path => !string.IsNullOrWhiteSpace(path));
    }

    private static string? FindExcelPath(string oneDriveRoot)
    {
        return Directory.EnumerateFiles(oneDriveRoot, "*ToDo*.xlsx", SearchOption.AllDirectories)
            .Where(static path =>
                !path.Contains("_test", StringComparison.OrdinalIgnoreCase) &&
                !path.Contains("_preview", StringComparison.OrdinalIgnoreCase) &&
                !path.Contains(".backup", StringComparison.OrdinalIgnoreCase) &&
                !path.Contains(".bak", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .FirstOrDefault();
    }

    private static string? TryReadGitConfig(string repositoryPath, string key)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "git",
                WorkingDirectory = repositoryPath,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            startInfo.ArgumentList.Add("config");
            startInfo.ArgumentList.Add(key);

            using var process = Process.Start(startInfo);
            if (process is null)
            {
                return null;
            }

            var output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit(5000);
            return process.ExitCode == 0 ? output : null;
        }
        catch
        {
            return null;
        }
    }
}
