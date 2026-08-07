using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Tests.Helpers;

internal sealed class NoOpRowCountMismatchAlertSender : IRowCountMismatchAlertSender
{
    public static NoOpRowCountMismatchAlertSender Instance { get; } = new();

    public Task SendRowCountMismatchAlertAsync(
        ReportSlotContext slot,
        IReadOnlyList<(string SourceLabel, int RowCount)> rowCounts,
        CancellationToken cancellationToken = default) =>
        Task.CompletedTask;
}
