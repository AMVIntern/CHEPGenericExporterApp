using System.Text;
using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Services.Email;

public sealed class MissingFileAlertSender : IMissingFileAlertSender
{
    private readonly IEmailSender _emailSender;
    private readonly EmailOptions _email;
    private readonly ILogger<MissingFileAlertSender> _logger;

    public MissingFileAlertSender(
        IEmailSender emailSender,
        IOptions<EmailOptions> emailOptions,
        ILogger<MissingFileAlertSender> logger)
    {
        _emailSender = emailSender;
        _email = emailOptions.Value;
        _logger = logger;
    }

    public async Task SendMissingFilesAlertAsync(
        IReadOnlyList<string> missingDescriptions,
        CancellationToken cancellationToken = default,
        ReportSlotContext? scheduledSlot = null)
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

        var subject = scheduledSlot is { } slot
            ? $"Missing some files – Shift {slot.Shift}, {slot.ReportDateDdMmmYyyy}"
            : (lines.Count == 1
                ? $"File is missing – {TruncateSubjectDetail(lines[0])}"
                : $"File is missing – {lines.Count} issues");

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
            _logger.LogInformation("Missing-file alert email sent to InternalAmvTeam ({Count} issue(s)).", lines.Count);
        else
            _logger.LogWarning("Missing-file alert email was not sent successfully.");
    }

    private static string TruncateSubjectDetail(string detail, int maxLen = 80)
    {
        if (string.IsNullOrEmpty(detail) || detail.Length <= maxLen)
            return detail;
        return detail[..(maxLen - 1)] + "…";
    }
}
