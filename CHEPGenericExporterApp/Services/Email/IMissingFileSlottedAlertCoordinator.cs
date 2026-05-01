using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Services.Email;

/// <summary>
/// When a <see cref="BeginSlottedBatch"/> scope is active, slotted missing-file lines for the same
/// slot are accumulated and sent as a single email when the scope is disposed (e.g. full pipeline run).
/// </summary>
public interface IMissingFileSlottedAlertCoordinator
{
    /// <summary>Starts accumulating slotted alerts until the returned scope is disposed asynchronously.</summary>
    IAsyncDisposable BeginSlottedBatch(
        ReportSlotContext slot,
        bool applyPerSlotMissingAlertLimit,
        CancellationToken cancellationToken);

    /// <summary>
    /// If a matching batch is active, appends lines and returns without sending. Otherwise sends immediately
    /// via <see cref="IMissingFileAlertSender"/>.
    /// </summary>
    Task SendOrEnqueueSlottedAsync(
        ReportSlotContext slot,
        bool applyPerSlotMissingAlertLimit,
        IReadOnlyList<string> lines,
        CancellationToken cancellationToken);
}
