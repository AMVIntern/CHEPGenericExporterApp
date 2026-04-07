using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Services.Email;

/// <summary>
/// Wraps SMTP sending and enqueues failed sends for background retry.
/// </summary>
public sealed class ReliableEmailSender : IEmailSender
{
    private readonly SmtpEmailSender _smtpSender;
    private readonly IEmailRetryQueue _retryQueue;
    private readonly SmtpOptions _smtpOptions;
    private readonly ILogger<ReliableEmailSender> _logger;

    public ReliableEmailSender(
        SmtpEmailSender smtpSender,
        IEmailRetryQueue retryQueue,
        IOptions<SmtpOptions> smtpOptions,
        ILogger<ReliableEmailSender> logger)
    {
        _smtpSender = smtpSender;
        _retryQueue = retryQueue;
        _smtpOptions = smtpOptions.Value;
        _logger = logger;
    }

    public async Task<bool> SendAsync(OutgoingEmail message, CancellationToken cancellationToken = default)
    {
        var sent = await _smtpSender.SendAsync(message, cancellationToken).ConfigureAwait(false);
        if (sent || !_smtpOptions.EnableBackgroundRetryQueue)
            return sent;

        _retryQueue.Enqueue(message);
        _logger.LogWarning(
            "Email send failed; queued for background retry in {Minutes} minute intervals. Queue size: {Count}. Subject: {Subject}",
            Math.Max(1, _smtpOptions.BackgroundRetryIntervalMinutes),
            _retryQueue.Count,
            message.Subject);
        return false;
    }
}
