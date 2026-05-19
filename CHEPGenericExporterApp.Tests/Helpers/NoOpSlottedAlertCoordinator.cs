using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Tests.Helpers;

internal sealed class NoOpSlottedAlertCoordinator : IMissingFileSlottedAlertCoordinator
{
    public static NoOpSlottedAlertCoordinator Instance { get; } = new();

    public ValueTask DisposeAsync() => ValueTask.CompletedTask;

    public IAsyncDisposable BeginSlottedBatch(
        ReportSlotContext slot,
        bool applyPerSlotMissingAlertLimit,
        CancellationToken cancellationToken) =>
        new BatchScope();

    public Task SendOrEnqueueSlottedAsync(
        ReportSlotContext slot,
        bool applyPerSlotMissingAlertLimit,
        IReadOnlyList<string> lines,
        CancellationToken cancellationToken) =>
        Task.CompletedTask;

    private sealed class BatchScope : IAsyncDisposable
    {
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
