using System.Diagnostics;
using System.Text;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Models;
using DBHWeeklyReport.Core.Services;

namespace DBHWeeklyReport.Infrastructure;

public sealed class GitCommitCollector : ICommitCollector
{
    private const string CommitMarker = "__COMMIT__";
    private const string BodyMarker = "__BODY__";
    private const string FilesMarker = "__FILES__";

    public async Task<(string? AuthorName, string? AuthorEmail, IReadOnlyList<CommitInfo> Commits)> CollectAsync(
        string repositoryPath,
        DateOnly startDate,
        DateOnly endDate,
        string? authorName,
        string? authorEmail,
        CancellationToken cancellationToken = default)
    {
        authorName = string.IsNullOrWhiteSpace(authorName) ? await RunGitAsync(repositoryPath, "config", "user.name") : authorName;
        authorEmail = string.IsNullOrWhiteSpace(authorEmail) ? await RunGitAsync(repositoryPath, "config", "user.email") : authorEmail;

        var args = new List<string>
        {
            "log",
            $"--since={startDate:yyyy-MM-dd} 00:00:00",
            $"--until={endDate:yyyy-MM-dd} 23:59:59",
            "--date=short",
            "--name-only",
            $"--pretty=format:{CommitMarker}%n%ad%x1f%an%x1f%ae%x1f%s%n{BodyMarker}%n%b%n{FilesMarker}",
        };

        if (!string.IsNullOrWhiteSpace(authorName))
        {
            args.Add($"--author={authorName}");
        }

        var output = await RunGitAsync(repositoryPath, args.ToArray(), cancellationToken);
        return (authorName?.Trim(), authorEmail?.Trim(), ParseCommits(output, authorEmail));
    }

    private static List<CommitInfo> ParseCommits(string output, string? authorEmail)
    {
        var commits = new List<CommitInfo>();
        string[]? header = null;
        var currentBodyLines = new List<string>();
        var currentFiles = new List<string>();
        var inBody = false;
        var inFiles = false;

        foreach (var line in output.Split(['\r', '\n'], StringSplitOptions.None))
        {
            if (line == CommitMarker)
            {
                TryAddCommit(commits, header, currentBodyLines, currentFiles, authorEmail);
                header = null;
                currentBodyLines = [];
                currentFiles = [];
                inBody = false;
                inFiles = false;
                continue;
            }

            if (header is null)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                header = line.Split('\x1f', 4);
                continue;
            }

            if (line == BodyMarker)
            {
                inBody = true;
                inFiles = false;
                continue;
            }

            if (line == FilesMarker)
            {
                inBody = false;
                inFiles = true;
                continue;
            }

            if (inBody)
            {
                currentBodyLines.Add(line.TrimEnd());
                continue;
            }

            if (inFiles && !string.IsNullOrWhiteSpace(line))
            {
                currentFiles.Add(line.Trim());
            }
        }

        TryAddCommit(commits, header, currentBodyLines, currentFiles, authorEmail);
        return commits;
    }

    private static void TryAddCommit(
        List<CommitInfo> commits,
        string[]? header,
        List<string> bodyLines,
        List<string> files,
        string? authorEmail)
    {
        if (header is null || header.Length < 4)
        {
            return;
        }

        var rawDate = header[0];
        var commitEmail = header[2];
        var subject = header[3].Trim();

        if (!string.IsNullOrWhiteSpace(authorEmail) &&
            !string.Equals(commitEmail.Trim(), authorEmail.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (WeeklyReportComposer.ShouldSkipSubject(subject))
        {
            return;
        }

        commits.Add(new CommitInfo(DateOnly.Parse(rawDate), subject, bodyLines.ToArray(), files.ToArray()));
    }

    private static async Task<string> RunGitAsync(string repositoryPath, params string[] arguments)
        => await RunGitAsync(repositoryPath, arguments, CancellationToken.None);

    private static async Task<string> RunGitAsync(string repositoryPath, string[] arguments, CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "git",
            WorkingDirectory = repositoryPath,
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
        var stdout = await process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderr = await process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"git {string.Join(' ', arguments)} failed with exit code {process.ExitCode}: {stderr}");
        }

        return stdout.Trim();
    }
}
