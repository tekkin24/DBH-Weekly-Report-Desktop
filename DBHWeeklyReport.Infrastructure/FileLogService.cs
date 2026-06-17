using DBHWeeklyReport.Core;

namespace DBHWeeklyReport.Infrastructure;

public sealed class FileLogService(Func<string> directoryAccessor) : ILogService
{
    public string LogDirectory => directoryAccessor();

    public async Task<string> WriteAsync(string category, string content, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(LogDirectory);
        var filePath = Path.Combine(LogDirectory, $"{DateTime.Now:yyyyMMdd-HHmmss}-{category}.log");
        await File.WriteAllTextAsync(filePath, content, cancellationToken);
        return filePath;
    }
}
