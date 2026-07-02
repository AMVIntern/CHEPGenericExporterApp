namespace CHEPGenericExporterApp.Models;

public sealed class ReportEvent
{
    public DateTimeOffset Timestamp { get; init; }
    public string Site { get; init; } = "";
    public string Shift { get; init; } = "";
    public DateOnly ReportDate { get; init; }
    public string TimeZone { get; init; } = "";
    public ReportEventKind Kind { get; init; }
    public bool Sent { get; init; }
    public ReportOutcomeReason Reason { get; init; }
    public string ReasonDetail { get; init; } = "";
    public IReadOnlyList<string> MissingInputs { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> DummyStations { get; init; } = Array.Empty<string>();
}
