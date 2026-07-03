namespace CHEPGenericExporterApp.Models;

public sealed class CombinedReportResult
{
    public string? ExcelFilePath { get; set; }
    public string? NormalizedCsvPath { get; set; }
    public string? NormalizedZipPath { get; set; }

    /// <summary>Station codes (e.g. S1, S2) for which a <c>*_DUMMY.csv</c> was generated this run.</summary>
    public IReadOnlyList<string> DummyStationsUsed { get; init; } = Array.Empty<string>();

    /// <summary>Normalized report site code (e.g. AUH4) for operational alerts.</summary>
    public string SiteCode { get; init; } = "";

    /// <summary>Specific missing-input lines (Gocator CSV, station files, etc.) when the report could not be produced.</summary>
    public IReadOnlyList<string> MissingInputs { get; init; } = Array.Empty<string>();
}
