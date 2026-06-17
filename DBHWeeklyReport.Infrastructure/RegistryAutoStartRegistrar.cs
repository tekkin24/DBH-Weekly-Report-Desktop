using System.Runtime.Versioning;
using Microsoft.Win32;
using DBHWeeklyReport.Core;

namespace DBHWeeklyReport.Infrastructure;

[SupportedOSPlatform("windows")]
public sealed class RegistryAutoStartRegistrar : IAutoStartRegistrar
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "DBHWeeklyReportDesktop";
    private const string BackgroundArgument = "--background";

    public bool IsEnabled(string executablePath)
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath);
        var value = key?.GetValue(ValueName) as string;
        return string.Equals(value, BuildCommand(executablePath), StringComparison.OrdinalIgnoreCase);
    }

    public void SetEnabled(string executablePath, bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath);
        if (enabled)
        {
            key.SetValue(ValueName, BuildCommand(executablePath));
        }
        else
        {
            key.DeleteValue(ValueName, false);
        }
    }

    private static string BuildCommand(string path) => $"\"{path}\" {BackgroundArgument}";
}
