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
    public string NormalizedReportSiteCode { get; set; }

    /// <summary>
    /// When true, a missing S1/S2/S4/S5 shift CSV is replaced with a generated <c>*_DUMMY.csv</c> (zeros, headers copied from another file in that folder).
    /// </summary>
    public bool CreateDummyStationShiftCsvWhenMissing { get; set; }

    /// <summary>
    /// Sub-header values (row 2 of station CSVs) that trigger a station prefix on the attribute name.
    /// E.g. ["B1","B2","B3"] — when the sub-header matches, the station prefix is prepended.
    /// </summary>
    public List<string> NormalizedReportBoardSubHeaders { get; set; } = new();

    /// <summary>
    /// Maps station code (S1/S2/S4/S5) to the prefix prepended to attributes whose sub-header is in NormalizedReportBoardSubHeaders.
    /// E.g. { "S1": "TOP_", "S2": "TOP_", "S5": "BOT_" }
    /// </summary>
    public Dictionary<string, string> NormalizedReportStationPrefixes { get; set; } = new();
}
