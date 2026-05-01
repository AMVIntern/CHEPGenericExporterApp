using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Services.Email;

public sealed class MissingFileSlottedAlertCoordinator : IMissingFileSlottedAlertCoordinator
{
    private readonly IMissingFileAlertSender _sender;
    private readonly AsyncLocal<BatchState?> _active = new();

    public MissingFileSlottedAlertCoordinator(IMissingFileAlertSender sender) =>
        _sender = sender;

    public IAsyncDisposable BeginSlottedBatch(
        ReportSlotContext slot,
        bool applyPerSlotMissingAlertLimit,
        CancellationToken cancellationToken)
    {
        if (_active.Value != null)
            return NoOpScope.Instance;

        var state = new BatchState(slot, applyPerSlotMissingAlertLimit, cancellationToken, new List<string>());
        _active.Value = state;
        return new SlottedBatchScope(this, state);
    }

    public async Task SendOrEnqueueSlottedAsync(
        ReportSlotContext slot,
        bool applyPerSlotMissingAlertLimit,
        IReadOnlyList<string> lines,
        CancellationToken cancellationToken)
    {
        if (TryEnqueue(slot, applyPerSlotMissingAlertLimit, lines))
            return;

        await _sender.SendMissingFilesAlertAsync(
                lines,
                cancellationToken,
                scheduledSlot: slot,
                applyPerSlotMissingAlertLimit: applyPerSlotMissingAlertLimit)
            .ConfigureAwait(false);
    }

    private bool TryEnqueue(
        ReportSlotContext slot,
        bool applyPerSlotMissingAlertLimit,
        IEnumerable<string> lines)
    {
        var st = _active.Value;
        if (st == null || st.Apply != applyPerSlotMissingAlertLimit || st.Slot != slot)
            return false;

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;
            var t = line.Trim();
            if (!st.Lines.Contains(t, StringComparer.Ordinal))
                st.Lines.Add(t);
        }

        return true;
    }

    private async ValueTask FlushAsync(BatchState state)
    {
        if (state.Lines.Count == 0 || state.CancellationToken.IsCancellationRequested)
            return;

        await _sender.SendMissingFilesAlertAsync(
                state.Lines,
                state.CancellationToken,
                scheduledSlot: state.Slot,
                applyPerSlotMissingAlertLimit: state.Apply)
            .ConfigureAwait(false);
    }

    private sealed record BatchState(
        ReportSlotContext Slot,
        bool Apply,
        CancellationToken CancellationToken,
        List<string> Lines);

    private sealed class SlottedBatchScope : IAsyncDisposable
    {
        private readonly MissingFileSlottedAlertCoordinator _parent;
        private readonly BatchState _state;

        public SlottedBatchScope(MissingFileSlottedAlertCoordinator parent, BatchState state)
        {
            _parent = parent;
            _state = state;
        }

        public ValueTask DisposeAsync()
        {
            if (!ReferenceEquals(_parent._active.Value, _state))
                return ValueTask.CompletedTask;

            _parent._active.Value = null;
            return _parent.FlushAsync(_state);
        }
    }

    private sealed class NoOpScope : IAsyncDisposable
    {
        public static readonly NoOpScope Instance = new();
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
