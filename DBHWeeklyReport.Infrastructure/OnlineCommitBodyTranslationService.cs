using System.Net.Http.Json;
using System.Text.Json;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Infrastructure;

public sealed class OnlineCommitBodyTranslationService : ICommitBodyTranslationService
{
    private static readonly HttpClient HttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(20),
    };

    private readonly Dictionary<string, string> _cache = new(StringComparer.Ordinal);
    private readonly object _cacheLock = new();

    public async Task<IReadOnlyList<CommitInfo>> TranslateAsync(
        IReadOnlyList<CommitInfo> commits,
        CancellationToken cancellationToken = default)
    {
        if (commits.Count == 0)
        {
            return commits;
        }

        var translatedCommits = new List<CommitInfo>(commits.Count);
        foreach (var commit in commits)
        {
            var bodyLines = new List<string>(commit.BodyLines.Count);
            foreach (var line in commit.BodyLines)
            {
                bodyLines.Add(await TranslateLineAsync(line, cancellationToken));
            }

            translatedCommits.Add(commit with { BodyLines = bodyLines.ToArray() });
        }

        return translatedCommits;
    }

    private async Task<string> TranslateLineAsync(string line, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return line;
        }

        lock (_cacheLock)
        {
            if (_cache.TryGetValue(line, out var cached))
            {
                return cached;
            }
        }

        var trimmed = line.Trim();
        var markerMatch = System.Text.RegularExpressions.Regex.Match(
            trimmed,
            @"^(?<marker>[-*]|\d+[.)])\s+(?<content>.+)$");

        var marker = markerMatch.Success ? $"{markerMatch.Groups["marker"].Value} " : string.Empty;
        var content = markerMatch.Success ? markerMatch.Groups["content"].Value.Trim() : trimmed;
        var translatedContent = await TranslateContentAsync(content, cancellationToken);

        var translated = marker + translatedContent;
        lock (_cacheLock)
        {
            _cache[line] = translated;
        }

        return translated;
    }

    private static async Task<string> TranslateContentAsync(string content, CancellationToken cancellationToken)
    {
        try
        {
            var url =
                $"https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=ja&dt=t&q={Uri.EscapeDataString(content)}";
            using var response = await HttpClient.GetAsync(url, cancellationToken);
            response.EnsureSuccessStatusCode();

            await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
            if (document.RootElement.ValueKind != JsonValueKind.Array || document.RootElement.GetArrayLength() == 0)
            {
                return content;
            }

            var segments = document.RootElement[0];
            if (segments.ValueKind != JsonValueKind.Array)
            {
                return content;
            }

            var translated = string.Concat(
                segments.EnumerateArray()
                    .Select(static segment =>
                        segment.ValueKind == JsonValueKind.Array && segment.GetArrayLength() > 0
                            ? segment[0].GetString()
                            : string.Empty));

            return string.IsNullOrWhiteSpace(translated) ? content : translated.Trim();
        }
        catch
        {
            return content;
        }
    }
}
