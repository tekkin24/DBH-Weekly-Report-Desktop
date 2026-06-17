using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Core.Services;

public static class SettingsValidator
{
    public static ValidationResult Validate(AppSettings settings)
    {
        var result = new ValidationResult();

        if (string.IsNullOrWhiteSpace(settings.RepositoryPath))
        {
            result.Errors.Add("Repository path is required.");
        }
        else if (!Directory.Exists(settings.RepositoryPath))
        {
            result.Errors.Add("Repository path does not exist.");
        }

        if (string.IsNullOrWhiteSpace(settings.ExcelPath))
        {
            result.Errors.Add("Excel file path is required.");
        }
        else if (!File.Exists(settings.ExcelPath))
        {
            result.Errors.Add("Excel file path does not exist.");
        }

        if (string.IsNullOrWhiteSpace(settings.LogDirectory))
        {
            result.Errors.Add("Log directory is required.");
        }

        return result;
    }
}
