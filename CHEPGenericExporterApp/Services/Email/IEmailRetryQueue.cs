using CHEPGenericExporterApp.Models;

namespace CHEPGenericExporterApp.Services.Email;

public interface IEmailRetryQueue
{
    void Enqueue(OutgoingEmail message);
    IReadOnlyList<OutgoingEmail> DrainBatch(int maxItems);
    int Count { get; }
}
