using ClosedXML.Excel;
using DBHWeeklyReport.Core;
using DBHWeeklyReport.Core.Models;
using DBHWeeklyReport.Core.Services;

namespace DBHWeeklyReport.Infrastructure;

public sealed class ClosedXmlExcelReportWriter : IExcelReportWriter
{
    private const string SheetDateCell = "B5";
    private const int DataStartRow = 8;

    public Task<string> WriteAsync(string excelPath, WeeklyReportPreview preview, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var sourcePath = Path.GetFullPath(excelPath);
        var backupPath = Path.Combine(
            Path.GetDirectoryName(sourcePath)!,
            $"{Path.GetFileNameWithoutExtension(sourcePath)}.backup-{DateTime.Now:yyyyMMdd-HHmmss}{Path.GetExtension(sourcePath)}");

        File.Copy(sourcePath, backupPath, true);

        using var workbook = new XLWorkbook(sourcePath);
        var templateSheet = ChooseTemplateSheet(workbook);
        var completedFillColor = ResolveCompletedFill(templateSheet);

        foreach (var report in preview.Reports)
        {
            var ws = EnsureMonthSheet(workbook, templateSheet, report.ReportDate);
            WriteDailyReport(ws, report, completedFillColor);
        }

        foreach (var group in preview.Reports.GroupBy(static report => WeeklyReportComposer.ResolveSheetTitle(report.ReportDate)))
        {
            var ws = workbook.Worksheet(group.Key);
            MergeTaskSpans(ws, group.OrderBy(static report => report.ReportDate).ToList());
        }

        workbook.SaveAs(sourcePath);
        return Task.FromResult(sourcePath);
    }

    private static IXLWorksheet ChooseTemplateSheet(XLWorkbook workbook)
    {
        var preferred = workbook.Worksheets.FirstOrDefault(static ws => ws.Name.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase));
        if (preferred is not null)
        {
            return preferred;
        }

        return workbook.Worksheet(workbook.Worksheets.Count);
    }

    private static XLColor ResolveCompletedFill(IXLWorksheet templateSheet)
    {
        for (var row = DataStartRow; row < DataStartRow + 31; row++)
        {
            if (string.Equals(templateSheet.Cell($"C{row}").GetString(), "完了", StringComparison.Ordinal))
            {
                return templateSheet.Cell($"D{row}").Style.Fill.BackgroundColor;
            }
        }

        return XLColor.FromHtml("#92D050");
    }

    private static IXLWorksheet EnsureMonthSheet(XLWorkbook workbook, IXLWorksheet templateSheet, DateOnly targetDate)
    {
        var expectedTitle = WeeklyReportComposer.ResolveSheetTitle(targetDate);
        var existing = workbook.Worksheets.FirstOrDefault(ws =>
            string.Equals(ws.Name, expectedTitle, StringComparison.OrdinalIgnoreCase) ||
            TryGetMonthStart(ws, out var monthStart) && monthStart == new DateOnly(targetDate.Year, targetDate.Month, 1));

        if (existing is not null)
        {
            return existing;
        }

        var copied = templateSheet.CopyTo(expectedTitle);
        copied.Cell(SheetDateCell).Value = new DateTime(targetDate.Year, targetDate.Month, 1);
        return copied;
    }

    private static bool TryGetMonthStart(IXLWorksheet worksheet, out DateOnly monthStart)
    {
        monthStart = default;
        var value = worksheet.Cell(SheetDateCell).Value;
        if (value.TryConvert(out DateTime dateTime))
        {
            monthStart = DateOnly.FromDateTime(dateTime.Date);
            return true;
        }

        return false;
    }

    private static void WriteDailyReport(IXLWorksheet ws, DailyReportEntry report, XLColor completedFillColor)
    {
        var row = RowForDate(report.ReportDate);

        ws.Cell(row, 3).Value = report.Status;
        ws.Cell(row, 4).Value = report.Task;
        ws.Cell(row, 8).Value = report.Detail;
        ws.Cell(row, 13).Value = report.Memo;

        var taskRange = ws.Range(row, 4, row, 7);
        taskRange.Style.Fill.BackgroundColor = completedFillColor;
        taskRange.Style.Alignment.WrapText = true;
        taskRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

        foreach (var column in new[] { 4, 8, 13 })
        {
            var cell = ws.Cell(row, column);
            cell.Style.Alignment.WrapText = true;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        }

        var baseHeight = ws.RowHeight > 0 ? ws.RowHeight : 15d;
        var lineCount = EstimateLineCount(report.Task, report.Detail, report.Memo);
        ws.Row(row).Height = baseHeight * Math.Min(3, Math.Max(1, lineCount));
    }

    private static void MergeTaskSpans(IXLWorksheet ws, IReadOnlyList<DailyReportEntry> reports)
    {
        foreach (var mergedRange in ws.MergedRanges.ToList())
        {
            if (mergedRange.RangeAddress.FirstAddress.ColumnNumber == 4 &&
                mergedRange.RangeAddress.LastAddress.ColumnNumber == 7 &&
                reports.Any(report =>
                {
                    var row = RowForDate(report.ReportDate);
                    return row >= mergedRange.RangeAddress.FirstAddress.RowNumber &&
                           row <= mergedRange.RangeAddress.LastAddress.RowNumber;
                }))
            {
                mergedRange.Unmerge();
            }
        }

        if (reports.Count == 0)
        {
            return;
        }

        var spanStart = reports[0];
        var previous = reports[0];

        for (var i = 1; i < reports.Count; i++)
        {
            var current = reports[i];
            var consecutive = current.ReportDate.DayNumber == previous.ReportDate.DayNumber + 1;
            var sameTask = string.Equals(current.Task, previous.Task, StringComparison.Ordinal);

            if (consecutive && sameTask)
            {
                previous = current;
                continue;
            }

            MergeSpan(ws, spanStart, previous);
            spanStart = current;
            previous = current;
        }

        MergeSpan(ws, spanStart, previous);
    }

    private static void MergeSpan(IXLWorksheet ws, DailyReportEntry start, DailyReportEntry end)
    {
        var startRow = RowForDate(start.ReportDate);
        var endRow = RowForDate(end.ReportDate);
        ws.Range(startRow, 4, endRow, 7).Merge();
        ws.Cell(startRow, 4).Style.Alignment.WrapText = true;
        ws.Cell(startRow, 4).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
    }

    private static int RowForDate(DateOnly date) => DataStartRow + date.Day - 1;

    private static int EstimateLineCount(params string[] values)
    {
        var segments = values.Where(static value => !string.IsNullOrWhiteSpace(value)).ToArray();
        if (segments.Length == 0)
        {
            return 1;
        }

        var longest = segments.Max(static segment => segment.Length);
        var explicitBreaks = segments.Max(static segment => segment.Count(static ch => ch == '\n') + 1);
        var estimatedFromLength = Math.Max(1, longest / 28 + 1);
        return Math.Max(explicitBreaks, estimatedFromLength);
    }
}
