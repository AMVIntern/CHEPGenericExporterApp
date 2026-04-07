namespace CHEPGenericExporterApp.Models;

public sealed class CombinedReportResult
{
    public string? ExcelFilePath { get; set; }
    public string? NormalizedCsvPath { get; set; }
    public string? NormalizedZipPath { get; set; }
}
