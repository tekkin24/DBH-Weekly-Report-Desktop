namespace DBHWeeklyReport.Infrastructure;

public static class AppPaths
{
    public static readonly string AppDataRoot = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "DBH Weekly Report Desktop");

    public static readonly string SettingsPath = Path.Combine(AppDataRoot, "settings.json");

    public static readonly string DefaultLogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "DBH Weekly Report Logs",
        "desktop");
}
