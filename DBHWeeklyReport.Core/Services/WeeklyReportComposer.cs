using System.Text.RegularExpressions;
using DBHWeeklyReport.Core.Models;

namespace DBHWeeklyReport.Core.Services;

public sealed partial class WeeklyReportComposer : IWeeklyReportComposer
{
    private const string StatusValue = "完了";

    private static readonly Regex[] SkipSubjects =
    [
        new Regex("^ok$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
        new Regex("^merge pull request", RegexOptions.IgnoreCase | RegexOptions.Compiled),
        new Regex("^revert\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled),
    ];

    private static readonly (Regex Pattern, string Label)[] AreaPatterns =
    [
        (Rx("materialprinting|material printing|material print"), "MaterialPrinting画面"),
        (Rx("productinspectiongroupededitor|product inspection grouped editor"), "ProductInspectionGroupedEditor"),
        (Rx("productinspectiontable|product inspection table"), "ProductInspectionTable"),
        (Rx("dailyreportsprintfiscal"), "DailyReportsPrintFiscal"),
        (Rx("product inspection print|inspection print"), "inspection帳票"),
        (Rx("admin?tensiletests|tensiletests|tensile sidebar|tensile"), "tensile画面"),
        (Rx("pdfdocumentservice|pdfdocumentscontroller|audit zip|audit package zip"), "PDF/ZIP出力処理"),
        (Rx("pdf-viewer|pdf viewer|dbheadpdfviewerhost"), "PDF viewer"),
        (Rx("dbhbrowserhelpers|index\\.html|dbheadpdfbridge"), "共通スクリプト"),
        (Rx("app\\.css"), "共通スタイル"),
        (Rx("dailyreportsentry"), "DailyReportsEntry"),
    ];

    private static readonly (Regex Pattern, string Label)[] ObjectPatterns =
    [
        (Rx("dailyreportsprintfiscal"), "DailyReportsPrintFiscal"),
        (Rx("productinspectionsheetgrouper"), "inspection帳票のロットグループ化"),
        (Rx("production month picker"), "production month picker component"),
        (Rx("unsaved changes dialog"), "未保存変更ダイアログ"),
        (Rx("input workspace spacing"), "入力ワークスペースの間隔"),
        (Rx("inspection toolbar spacing"), "inspection toolbar の間隔"),
        (Rx("inspection toolbar"), "inspection toolbar"),
        (Rx("material pdf reload on company filter changes"), "会社フィルター変更時のMaterial PDF再読み込み"),
        (Rx("product inspection table form inline"), "ProductInspectionTableフォームのinline表示"),
        (Rx("product inspection sheets inline"), "inspection帳票のinline表示"),
        (Rx("pdf navigation buttons fixed during zoom"), "ズーム時のPDFナビゲーションボタン固定"),
        (Rx("pdf viewer toolbar actions on load"), "PDF viewer toolbar の初期表示"),
        (Rx("material print actions into pdf viewer toolbar"), "MaterialPrinting操作のPDF viewer toolbar への移動"),
        (Rx("material printing workspace"), "MaterialPrintingワークスペース"),
        (Rx("material printing studio layout"), "MaterialPrinting studio レイアウト"),
        (Rx("audit package zip download"), "audit ZIPダウンロード機能"),
        (Rx("audit zip browser download"), "audit ZIPのブラウザダウンロード"),
        (Rx("audit zip validation diagnostics"), "audit ZIP検証ログ"),
        (Rx("audit zip internal paths"), "audit ZIP内部パス"),
        (Rx("audit zip writer compile error"), "audit ZIP writer のコンパイルエラー"),
        (Rx("windows-compatible store entries"), "Windows互換のZIP格納形式"),
        (Rx("windows explorer"), "Windows Explorer対応"),
        (Rx("development api login endpoint"), "開発用API login endpoint"),
        (Rx("inspection dossier pdfs"), "inspection dossier PDF"),
        (Rx("inline inspection print loading stalls"), "inline inspection 印刷の停止問題"),
        (Rx("productinspectiongroupededitor|product inspection grouped editor"), "ProductInspectionGroupedEditor"),
        (Rx("tensile sidebar"), "tensile sidebar"),
        (Rx("all companies option"), "全会社オプション"),
        (Rx("hidden material printing ui"), "非表示のMaterialPrinting UI"),
        (Rx("dossier labels"), "dossier ラベル"),
        (Rx("dossier categories"), "dossier 分類"),
        (Rx("pdf export helper"), "PDF export helper"),
        (Rx("auth browser helper"), "auth browser helper"),
        (Rx("pdf actions and rename flow"), "PDF操作と名前変更フロー"),
        (Rx("load in productinspectiongroupededitor"), "ProductInspectionGroupedEditor読み込み"),
    ];

    private static readonly (Regex Pattern, string Verb, string Noun)[] ActionPatterns =
    [
        (Rx("\\b(add|create)\\b"), "作成する", "追加"),
        (Rx("\\b(remove|delete)\\b"), "削除する", "削除"),
        (Rx("\\b(separate|isolate|split)\\b"), "分離する", "分離"),
        (Rx("\\b(share|common)\\b"), "共通化する", "共通化"),
        (Rx("\\b(move)\\b"), "移動する", "移動"),
        (Rx("\\b(rename)\\b"), "変更する", "変更"),
        (Rx("\\b(normalize|unify|standardize)\\b"), "統一する", "統一"),
        (Rx("\\b(regenerate)\\b"), "再生成する", "再生成"),
        (Rx("\\b(redesign|restyle)\\b"), "再設計する", "再設計"),
        (Rx("\\b(adjust|tune)\\b"), "調整する", "調整"),
        (Rx("\\b(fix)\\b"), "修正する", "修正"),
        (Rx("\\b(harden|prevent|improve)\\b"), "改善する", "改善"),
        (Rx("\\b(update|match|keep|load|open|render|embed)\\b"), "更新する", "更新"),
    ];

    private static readonly (Regex Pattern, Func<Match, string> Builder)[] SubjectPatterns =
    [
        (Rx("enhance fiscal transform preview styles and structure in (.+)"), _ => "DailyReportsPrintFiscalの年度変換プレビューの構成と表示スタイルを改善する"),
        (Rx("refactor layout and styles for production month picker components"), _ => "生産月選択部品のレイアウトと表示スタイルを見直す"),
        (Rx("update parameter attributes .* improved query handling"), _ => "DailyReportsPrintFiscalの画面引数の受け渡し処理を改善する"),
        (Rx("update unsaved changes dialog button styles and reorder report type options"), _ => "未保存変更ダイアログのボタンスタイルを修正し、報告種別オプションの順序を調整する"),
        (Rx("simplify success messages across multiple pages to a consistent format"), _ => "複数画面の成功メッセージ表示を統一する"),
        (Rx("refactor pdf generation logic in (.+) and enhance document queuing process"), match => $"{DetectObject(match.Groups[1].Value, [])}のPDF生成処理を見直し、文書登録処理を改善する"),
        (Rx("implement unsaved changes prompt dialog and state management across multiple pages"), _ => "複数画面に未保存変更確認ダイアログと状態管理を追加する"),
        (Rx("enhance (.+) with improved button states and keyboard handling"), match => $"{DetectObject(match.Groups[1].Value, [])}のボタン状態とキーボード操作を改善する"),
        (Rx("add display functions for quantity and dash values in (.+)"), match => $"{DetectObject(match.Groups[1].Value, [])}に数量値と記号表示の処理を追加する"),
        (Rx("improve lot grouping logic in (.+)"), _ => "inspection帳票のロットグループ化ロジックを改善する"),
        (Rx("enhance product inspection features with dbringshipmentbox integration"), _ => "ProductInspectionにDbRingShipmentBox連携機能を追加する"),
    ];

    public WeeklyReportPreview Compose(
        string repositoryPath,
        string excelPath,
        string? authorName,
        string? authorEmail,
        DateOnly referenceDate,
        IReadOnlyList<CommitInfo> commits)
    {
        var window = ResolveWeekWindow(referenceDate);
        var reports = BuildDailyReports(window.WeekStart, window.WeekEnd, commits);

        return new WeeklyReportPreview
        {
            RepositoryPath = repositoryPath,
            ExcelPath = excelPath,
            AuthorName = authorName,
            AuthorEmail = authorEmail,
            WeekStart = window.WeekStart,
            WeekEnd = window.WeekEnd,
            CommitCount = commits.Count,
            Reports = reports,
            TouchedSheets = reports.Select(static r => ResolveSheetTitle(r.ReportDate)).Distinct().OrderBy(static r => r).ToArray(),
        };
    }

    public static WeekWindow ResolveWeekWindow(DateOnly referenceDate)
    {
        var diff = ((int)referenceDate.DayOfWeek + 6) % 7;
        var monday = referenceDate.AddDays(-diff);
        var friday = monday.AddDays(4);
        return new WeekWindow
        {
            WeekStart = monday,
            WeekEnd = referenceDate.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday
                ? friday
                : (referenceDate < friday ? referenceDate : friday),
        };
    }

    public static string ResolveSheetTitle(DateOnly targetDate) => new DateOnly(targetDate.Year, targetDate.Month, 1).ToString("dd-MM-yyyy");

    public static bool ShouldSkipSubject(string subject) => SkipSubjects.Any(pattern => pattern.IsMatch(subject.Trim()));

    private static List<DailyReportEntry> BuildDailyReports(DateOnly startDate, DateOnly endDate, IReadOnlyList<CommitInfo> commits)
    {
        var byDay = commits.GroupBy(static c => c.CommitDate).ToDictionary(static g => g.Key, static g => g.ToList());
        var result = new List<DailyReportEntry>();

        for (var current = startDate; current <= endDate; current = current.AddDays(1))
        {
            if (!byDay.TryGetValue(current, out var dayCommits) || dayCommits.Count == 0)
            {
                continue;
            }

            result.Add(BuildDailyReport(current, dayCommits));
        }

        return result;
    }

    private static DailyReportEntry BuildDailyReport(DateOnly reportDate, List<CommitInfo> commits)
    {
        var dominantArea = commits
            .Select(commit => DetectArea([commit.Subject, .. commit.Files]))
            .GroupBy(static label => label)
            .OrderByDescending(static g => g.Count())
            .First()
            .Key;

        var actionGroups = commits
            .Select(commit => DetectAction(commit.Subject).Verb)
            .GroupBy(static verb => verb)
            .OrderByDescending(static g => g.Count())
            .ToList();

        var dominantAction = actionGroups.Count > 1 ? "更新する" : actionGroups[0].Key;

        var translated = new List<string>();
        foreach (var commit in commits)
        {
            var text = TranslateSubject(commit.Subject, commit.Files);
            if (translated.Contains(text, StringComparer.Ordinal))
            {
                continue;
            }

            translated.Add(text);
            if (translated.Count == 6)
            {
                break;
            }
        }

        return new DailyReportEntry
        {
            ReportDate = reportDate,
            Status = StatusValue,
            Task = $"{dominantArea}を{dominantAction}",
            Detail = BuildDetail(commits),
            Memo = string.Join('。', translated),
            CommitCount = commits.Count,
        };
    }

    private static string BuildDetail(IEnumerable<CommitInfo> commits)
    {
        var details = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var commit in commits)
        {
            var label = DetectObject(commit.Subject, commit.Files);
            if (!seen.Add(label))
            {
                continue;
            }

            details.Add(label);
            if (details.Count == 3)
            {
                break;
            }
        }

        if (details.Count > 0)
        {
            return string.Join('、', details);
        }

        return string.Join('、', commits
            .Select(commit => DetectArea([commit.Subject, .. commit.Files]))
            .GroupBy(static label => label)
            .OrderByDescending(static g => g.Count())
            .Take(2)
            .Select(static g => g.Key));
    }

    private static string DetectArea(IEnumerable<string> texts)
    {
        var combined = string.Join(' ', texts);
        foreach (var (pattern, label) in AreaPatterns)
        {
            if (pattern.IsMatch(combined))
            {
                return label;
            }
        }

        return "共通機能";
    }

    private static (string Verb, string Noun) DetectAction(string subject)
    {
        foreach (var (pattern, verb, noun) in ActionPatterns)
        {
            if (pattern.IsMatch(subject))
            {
                return (verb, noun);
            }
        }

        return ("更新する", "更新");
    }

    private static string DetectObject(string subject, IReadOnlyList<string> files)
    {
        foreach (var (pattern, label) in ObjectPatterns)
        {
            if (pattern.IsMatch(subject))
            {
                return label;
            }
        }

        return DetectArea([subject, .. files]);
    }

    private static string TranslateSubject(string subject, IReadOnlyList<string> files)
    {
        foreach (var (pattern, builder) in SubjectPatterns)
        {
            var match = pattern.Match(subject.Trim());
            if (match.Success)
            {
                return builder(match);
            }
        }

        var (verb, noun) = DetectAction(subject);
        var target = DetectObject(subject, files);

        return noun switch
        {
            "移動" when target.EndsWith("移動", StringComparison.Ordinal) => $"{target}を対応する",
            "再設計" => $"{target}を再設計する",
            "再生成" => $"{target}を再生成する",
            "共通化" => $"{target}を共通化する",
            "統一" => $"{target}を統一する",
            "削除" => $"{target}を削除する",
            "追加" => $"{target}を作成する",
            "分離" => $"{target}を分離する",
            "変更" => $"{target}を変更する",
            "移動" => $"{target}を移動する",
            "調整" => $"{target}を調整する",
            "修正" => $"{target}を修正する",
            "改善" => $"{target}を改善する",
            _ => $"{target}を{verb}",
        };
    }

    private static Regex Rx(string pattern) => new(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
}
