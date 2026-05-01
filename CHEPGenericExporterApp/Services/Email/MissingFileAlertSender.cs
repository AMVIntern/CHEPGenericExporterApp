using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Services.Email;

public sealed class MissingFileAlertSender : IMissingFileAlertSender
{
    /// <summary>Serializes check → send → record for the same shift/date so scheduler and recovery cannot double-send.</summary>
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> PerSlotMissingAlertLocks =
        new(StringComparer.Ordinal);

    private readonly IEmailSender _emailSender;
    private readonly EmailOptions _email;
    private readonly CsvAuditLogger _csvAuditLogger;
    private readonly ILogger<MissingFileAlertSender> _logger;

    public MissingFileAlertSender(
        IEmailSender emailSender,
        IOptions<EmailOptions> emailOptions,
        CsvAuditLogger csvAuditLogger,
        ILogger<MissingFileAlertSender> logger)
    {
        _emailSender = emailSender;
        _email = emailOptions.Value;
        _csvAuditLogger = csvAuditLogger;
        _logger = logger;
    }

    public async Task SendMissingFilesAlertAsync(
        IReadOnlyList<string> missingDescriptions,
        CancellationToken cancellationToken = default,
        ReportSlotContext? scheduledSlot = null,
        bool applyPerSlotMissingAlertLimit = false)
    {
        var lines = missingDescriptions.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).ToList();
        if (lines.Count == 0)
            return;

        if (_email.InternalAmvTeam == null || _email.InternalAmvTeam.Count == 0)
        {
            _logger.LogDebug("InternalAmvTeam is not configured; skipping missing-file alert.");
            return;
        }

        if (string.IsNullOrWhiteSpace(_email.FromAddress))
        {
            _logger.LogWarning("Email FromAddress is not configured; cannot send missing-file alert.");
            return;
        }

        var maxMissingAlerts = _email.MaxMissingFileAlertsPerShiftDate;
        var usePerSlotLimit = applyPerSlotMissingAlertLimit && scheduledSlot is { } && maxMissingAlerts > 0;
        SemaphoreSlim? slotGate = null;
        if (usePerSlotLimit && scheduledSlot is { } gateSlot)
        {
            slotGate = PerSlotMissingAlertLocks.GetOrAdd(
                PerSlotLockKey(gateSlot.Shift, gateSlot.ReportDate),
                _ => new SemaphoreSlim(1, 1));
        }

        if (slotGate != null)
            await slotGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (usePerSlotLimit && scheduledSlot is { } limitSlot)
            {
                _csvAuditLogger.EnsureRow(limitSlot.Shift, limitSlot.ReportDate);
                if (!_csvAuditLogger.CanSendMissingFileAlert(limitSlot.Shift, limitSlot.ReportDate, maxMissingAlerts))
                {
                    _logger.LogInformation(
                        "Missing-file alert suppressed (limit {Max} reached or finalized) for Shift {Shift}, Date {Date}.",
                        maxMissingAlerts,
                        limitSlot.Shift,
                        limitSlot.ReportDateDdMmmYyyy);
                    return;
                }
            }

            var subject = scheduledSlot is { } slot
                ? FormatWithSlotTemplate(_email.MissingFileAlertSubjectWithSlotTemplate, slot.Shift, slot.ReportDateDdMmmYyyy)
                : (lines.Count == 1
                    ? FormatSingleIssueSubject(_email.MissingFileAlertSubjectSingleIssueTemplate, lines[0])
                    : FormatMultiIssueSubject(_email.MissingFileAlertSubjectMultiIssueTemplate, lines.Count));

            var body = new StringBuilder();
            foreach (var line in lines)
                body.AppendLine($"• {line}");
            body.AppendLine();
            //body.AppendLine($"Machine: {Environment.MachineName}");
            //body.AppendLine($"Time (UTC): {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss}Z");

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
            {
                _logger.LogInformation("Missing-file alert email sent to InternalAmvTeam ({Count} issue(s)).", lines.Count);
                if (usePerSlotLimit && scheduledSlot is { } recordedSlot)
                    _csvAuditLogger.RecordMissingFileAlertSent(recordedSlot.Shift, recordedSlot.ReportDate, maxMissingAlerts);
            }
            else
                _logger.LogWarning("Missing-file alert email was not sent successfully.");
        }
        finally
        {
            slotGate?.Release();
        }
    }

    private static string PerSlotLockKey(string shift, DateOnly date) =>
        string.Create(CultureInfo.InvariantCulture, $"{shift.Trim()}:{date:yyyy-MM-dd}");

    private static string FormatWithSlotTemplate(string template, string shift, string date)
    {
        var t = string.IsNullOrWhiteSpace(template)
            ? "Missing some files – Shift {shift}, {date}"
            : template;
        return t
            .Replace("{shift}", shift.Trim(), StringComparison.Ordinal)
            .Replace("{date}", date.Trim(), StringComparison.Ordinal);
    }

    private static string FormatSingleIssueSubject(string template, string detailLine)
    {
        var detail = TruncateSubjectDetail(detailLine);
        var t = string.IsNullOrWhiteSpace(template)
            ? "File is missing – {detail}"
            : template;
        return t.Replace("{detail}", detail, StringComparison.Ordinal);
    }

    private static string FormatMultiIssueSubject(string template, int count)
    {
        var t = string.IsNullOrWhiteSpace(template)
            ? "File is missing – {count} issues"
            : template;
        return t.Replace("{count}", count.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    private static string TruncateSubjectDetail(string detail, int maxLen = 80)
    {
        if (string.IsNullOrEmpty(detail) || detail.Length <= maxLen)
            return detail;
        return detail[..(maxLen - 1)] + "…";
    }
}
