using CHEPGenericExporterApp.Models;

namespace CHEPGenericExporterApp.Services.Email;

public interface IEmailSender
{
    Task<bool> SendAsync(OutgoingEmail message, CancellationToken cancellationToken = default);
}
