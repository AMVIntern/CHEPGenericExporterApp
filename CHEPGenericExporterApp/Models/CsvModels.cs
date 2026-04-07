namespace CHEPGenericExporterApp.Models;

public sealed class CsvData
{
    public string[] Headers { get; set; } = Array.Empty<string>();
    public List<CsvRow> Rows { get; set; } = new();
}

public sealed class CsvRow
{
    public Dictionary<string, string> Data { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public DateTime? FullTimestamp { get; set; }
}
