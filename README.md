# DBH Weekly Report Desktop

Desktop app for previewing and writing weekly work reports from local git commits into the Excel workbook.

## Current behavior

- Runs as a Windows tray app.
- Can start automatically with Windows.
- Default schedule: every Friday at `16:00`.
- Generates a preview first, then writes to Excel only after confirmation.
- Uses the existing proven Python Excel writer for workbook compatibility.

## First run

1. Open `DBHWeeklyReport.App.exe`.
2. Check `Repository path` and `Excel file path`.
3. Click `Save`.
4. Leave the app running in tray.

The app tries to auto-detect:

- `DBH.slnx`
- the latest `*ToDo*.xlsx`
- git `user.name`
- git `user.email`

## Important paths

- Settings:
  - `%AppData%\\DBH Weekly Report Desktop\\settings.json`
- Logs:
  - `%UserProfile%\\Documents\\DBH Weekly Report Logs\\desktop`
- Existing Python writer:
  - `%UserProfile%\\Documents\\DBH Weekly Report Automation\\weekly_report\\fill_weekly_report.py`

## Build

```powershell
dotnet build .\\DBHWeeklyReportDesktop.slnx -c Release
```

## Publish

```powershell
dotnet publish .\\DBHWeeklyReport.App\\DBHWeeklyReport.App.csproj -c Release -r win-x64 --self-contained true
```
