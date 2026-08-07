using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Services.Email;

/// <summary>Sends an internal alert when the row counts across Gocator Top, Gocator Bottom, and stations 1/2/4/5
/// spread too far apart.</summary>
public interface IRowCountMismatchAlertSender
{
    /// <param name="rowCounts">Row count read directly from each raw input file for this shift (Gocator Top,
    /// Gocator Bottom, then each station), in display order. The sender identifies the lowest and highest counts
    /// to build the alert.</param>
    Task SendRowCountMismatchAlertAsync(
        ReportSlotContext slot,
        IReadOnlyList<(string SourceLabel, int RowCount)> rowCounts,
        CancellationToken cancellationToken = default);
}
