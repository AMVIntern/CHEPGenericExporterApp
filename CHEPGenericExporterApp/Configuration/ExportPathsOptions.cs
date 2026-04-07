namespace CHEPGenericExporterApp.Configuration;

/// <summary>Input/output folders for CSV merge and Excel generation. Paths may be absolute or relative to the app base directory.</summary>
public sealed class ExportPathsOptions
{
    public const string SectionName = "ExportPaths";

    public string TopCsvFolder { get; set; } = "Csvs/Top";
    public string BottomCsvFolder { get; set; } = "Csvs/Bottom";
    public string GocatorCombinedFolder { get; set; } = "Csvs/Combined";
    public string S1Folder { get; set; } = "Csvs/S1_DLDataLogs";
    public string S2Folder { get; set; } = "Csvs/S2_DLDataLogs";
    public string S4Folder { get; set; } = "Csvs/Station4Defects";
    public string S5Folder { get; set; } = "Csvs/Station5Defects";
    public string CombinedReportOutputFolder { get; set; } = "Csvs/Combined";

    /// <summary>Site column in normalized Power BI CSV.</summary>
    public string NormalizedReportSiteCode { get; set; } = "AUB6";
}
