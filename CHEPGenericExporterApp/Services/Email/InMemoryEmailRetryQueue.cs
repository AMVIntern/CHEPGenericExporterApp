using System.Collections.Concurrent;
using CHEPGenericExporterApp.Models;

namespace CHEPGenericExporterApp.Services.Email;

public sealed class InMemoryEmailRetryQueue : IEmailRetryQueue
{
    private readonly ConcurrentQueue<OutgoingEmail> _queue = new();
    private int _count;

    public int Count => Volatile.Read(ref _count);

    public void Enqueue(OutgoingEmail message)
    {
        _queue.Enqueue(Clone(message));
        Interlocked.Increment(ref _count);
    }

    public IReadOnlyList<OutgoingEmail> DrainBatch(int maxItems)
    {
        var items = new List<OutgoingEmail>(Math.Max(1, maxItems));
        while (items.Count < maxItems && _queue.TryDequeue(out var item))
        {
            items.Add(item);
            Interlocked.Decrement(ref _count);
        }

        return items;
    }

    private static OutgoingEmail Clone(OutgoingEmail src) =>
        new()
        {
            From = src.From,
            To = src.To.ToList(),
            Cc = src.Cc?.ToList(),
            Subject = src.Subject,
            Body = src.Body,
            PrimaryAttachmentPath = src.PrimaryAttachmentPath,
            AdditionalAttachmentPaths = src.AdditionalAttachmentPaths?.ToList()
        };
}
