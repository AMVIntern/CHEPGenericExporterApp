using System.Text;
using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Services.Email;

/// <summary>Internal-only alert: fires when the spread between the lowest and highest row count among Gocator
/// and stations 1/2/4/5 exceeds <see cref="EmailOptions.RowCountMismatchThreshold"/> rows.</summary>
public sealed class RowCountMismatchAlertSender : IRowCountMismatchAlertSender
{
    private readonly IEmailSender _emailSender;
    private readonly EmailOptions _email;
    private readonly ILogger<RowCountMismatchAlertSender> _logger;

    public RowCountMismatchAlertSender(
        IEmailSender emailSender,
        IOptions<EmailOptions> emailOptions,
        ILogger<RowCountMismatchAlertSender> logger)
    {
        _emailSender = emailSender;
        _email = emailOptions.Value;
        _logger = logger;
    }

    public async Task SendRowCountMismatchAlertAsync(
        ReportSlotContext slot,
        IReadOnlyList<(string SourceLabel, int RowCount)> rowCounts,
        CancellationToken cancellationToken = default)
    {
        if (rowCounts.Count == 0)
            return;

        if (_email.InternalAmvTeam == null || _email.InternalAmvTeam.Count == 0)
        {
            _logger.LogDebug("InternalAmvTeam is not configured; skipping row-count-mismatch alert.");
            return;
        }

        if (string.IsNullOrWhiteSpace(_email.FromAddress))
        {
            _logger.LogWarning("Email FromAddress is not configured; cannot send row-count-mismatch alert.");
            return;
        }

        var min = rowCounts.MinBy(r => r.RowCount);
        var max = rowCounts.MaxBy(r => r.RowCount);
        var diff = max.RowCount - min.RowCount;

        var subject = (string.IsNullOrWhiteSpace(_email.RowCountMismatchAlertSubjectTemplate)
                ? "Row count mismatch – Shift {shift}, {date}"
                : _email.RowCountMismatchAlertSubjectTemplate)
            .Replace("{shift}", slot.Shift, StringComparison.Ordinal)
            .Replace("{date}", slot.ReportDateDdMmmYyyy, StringComparison.Ordinal);

        var summaryTemplate = string.IsNullOrWhiteSpace(_email.RowCountMismatchAlertSummaryTemplate)
            ? "Row count difference of {diff} for Shift {shift}, Date {date}: lowest is {minSource} ({minCount} rows), highest is {maxSource} ({maxCount} rows)."
            : _email.RowCountMismatchAlertSummaryTemplate;

        var summary = summaryTemplate
            .Replace("{diff}", diff.ToString(), StringComparison.Ordinal)
            .Replace("{shift}", slot.Shift, StringComparison.Ordinal)
            .Replace("{date}", slot.ReportDateDdMmmYyyy, StringComparison.Ordinal)
            .Replace("{minSource}", min.SourceLabel, StringComparison.Ordinal)
            .Replace("{minCount}", min.RowCount.ToString(), StringComparison.Ordinal)
            .Replace("{maxSource}", max.SourceLabel, StringComparison.Ordinal)
            .Replace("{maxCount}", max.RowCount.ToString(), StringComparison.Ordinal);

        var lineTemplate = string.IsNullOrWhiteSpace(_email.RowCountMismatchAlertLineTemplate)
            ? "{source}: {count} rows"
            : _email.RowCountMismatchAlertLineTemplate;

        var body = new StringBuilder();
        body.AppendLine(summary);
        body.AppendLine();
        body.AppendLine("Row counts:");
        foreach (var r in rowCounts)
        {
            var line = lineTemplate
                .Replace("{source}", r.SourceLabel, StringComparison.Ordinal)
                .Replace("{count}", r.RowCount.ToString(), StringComparison.Ordinal)
                .Replace("{shift}", slot.Shift, StringComparison.Ordinal)
                .Replace("{date}", slot.ReportDateDdMmmYyyy, StringComparison.Ordinal);
            body.AppendLine($"• {line}");
        }

        var message = new OutgoingEmail
        {
            From = _email.FromAddress.Trim(),
            To = _email.InternalAmvTeam,
            Cc = null,
            Subject = subject,
            Body = body.ToString(),
            PrimaryAttachmentPath = null,
            AdditionalAttachmentPaths = null
        };

        var sent = await _emailSender.SendAsync(message, cancellationToken).ConfigureAwait(false);
        if (sent)
            _logger.LogInformation(
                "Row-count-mismatch alert email sent to InternalAmvTeam (diff {Diff}: {MinSource}={MinCount}, {MaxSource}={MaxCount}).",
                diff, min.SourceLabel, min.RowCount, max.SourceLabel, max.RowCount);
        else
            _logger.LogWarning("Row-count-mismatch alert email was not sent successfully.");
    }
}
