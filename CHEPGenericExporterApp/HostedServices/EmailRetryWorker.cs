using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Services.Email;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.HostedServices;

/// <summary>
/// Retries queued failed emails periodically (default every 10 minutes) until sent.
/// </summary>
public sealed class EmailRetryWorker : BackgroundService
{
    private readonly IEmailRetryQueue _retryQueue;
    private readonly SmtpEmailSender _smtpSender;
    private readonly IOptions<SmtpOptions> _smtpOptions;
    private readonly ILogger<EmailRetryWorker> _logger;

    public EmailRetryWorker(
        IEmailRetryQueue retryQueue,
        SmtpEmailSender smtpSender,
        IOptions<SmtpOptions> smtpOptions,
        ILogger<EmailRetryWorker> logger)
    {
        _retryQueue = retryQueue;
        _smtpSender = smtpSender;
        _smtpOptions = smtpOptions;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            var o = _smtpOptions.Value;
            var interval = TimeSpan.FromMinutes(Math.Clamp(o.BackgroundRetryIntervalMinutes, 1, 1440));

            if (!o.EnableBackgroundRetryQueue)
            {
                await Task.Delay(interval, stoppingToken).ConfigureAwait(false);
                continue;
            }

            await RetryBatchAsync(stoppingToken).ConfigureAwait(false);
            await Task.Delay(interval, stoppingToken).ConfigureAwait(false);
        }
    }

    private async Task RetryBatchAsync(CancellationToken cancellationToken)
    {
        var items = _retryQueue.DrainBatch(maxItems: 100);
        if (items.Count == 0)
            return;

        var sent = 0;
        var failed = 0;
        foreach (var mail in items)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var ok = await _smtpSender.SendAsync(mail, cancellationToken).ConfigureAwait(false);
            if (ok)
            {
                sent++;
            }
            else
            {
                failed++;
                _retryQueue.Enqueue(mail);
            }
        }

        _logger.LogInformation(
            "Background email retry batch completed. Sent: {Sent}, Requeued: {Requeued}, Pending: {Pending}.",
            sent,
            failed,
            _retryQueue.Count);
    }
}
