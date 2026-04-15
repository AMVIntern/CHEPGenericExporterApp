using System.Globalization;
using System.Text.RegularExpressions;
using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Reports;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Services;

/// <summary>Runs Gocator CSV merge, combined Excel generation, and optional email send.</summary>
public sealed class ExportPipeline
{
    private static readonly Regex ReportShiftDateInFileName = new(
        @"_Shift_(\d+)_(.+)$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private readonly GocatorCsvMergeService _gocatorMerge;
    private readonly CombinedExcelReportService _excelReport;
    private readonly IEmailSender _emailSender;
    private readonly IScheduleCalculator _scheduleCalculator;
    private readonly IMissingFileAlertSender _missingFileAlerts;
    private readonly EmailOptions _email;
    private readonly ILogger<ExportPipeline> _logger;

    public ExportPipeline(
        GocatorCsvMergeService gocatorMerge,
        CombinedExcelReportService excelReport,
        IEmailSender emailSender,
        IScheduleCalculator scheduleCalculator,
        IMissingFileAlertSender missingFileAlerts,
        IOptions<EmailOptions> emailOptions,
        ILogger<ExportPipeline> logger)
    {
        _gocatorMerge = gocatorMerge;
        _excelReport = excelReport;
        _emailSender = emailSender;
        _scheduleCalculator = scheduleCalculator;
        _missingFileAlerts = missingFileAlerts;
        _email = emailOptions.Value;
        _logger = logger;
    }

    /// <summary>Gocator merge for the scheduled slot, then email with the merged CSV (same idea as combined + email).</summary>
    public Task RunGocatorMergeOnlyAsync(ScheduledJob job, CancellationToken cancellationToken = default) =>
        RunGocatorMergeWithOptionalEmailAsync(job, sendEmailAfterMerge: true, cancellationToken);

    /// <summary>Combined Excel for the same scheduled shift/date as the Gocator step, then email.</summary>
    public async Task RunCombinedReportAndEmailAsync(
        ScheduledJob job,
        CancellationToken cancellationToken = default,
        ReportSlotContext? slotContext = null)
    {
        _logger.LogInformation("Running combined Excel report and email.");

        var ctx = slotContext ?? _scheduleCalculator.ResolveReportContext(job);
        var reportResult = await _excelReport.GenerateCombinedExcelReportAsync(ctx, cancellationToken).ConfigureAwait(false);

        if (reportResult == null ||
            string.IsNullOrEmpty(reportResult.ExcelFilePath) ||
            !File.Exists(reportResult.ExcelFilePath))
        {
            _logger.LogWarning("Combined Excel report was not produced; skipping email.");
            return;
        }

        if (!_email.BypassReportAttachmentSlotCheck &&
            !ReportFileMatchesScheduledSlot(reportResult.ExcelFilePath, ctx))
        {
            _logger.LogWarning(
                "Combined report output does not match scheduled shift/date; skipping customer email.");
            await _missingFileAlerts.SendMissingFilesAlertAsync(
                new[]
                {
                    $"Combined report is not for scheduled Shift {ctx.Shift}, Date {ctx.ReportDateDdMmmYyyy} (file: {Path.GetFileName(reportResult.ExcelFilePath)})."
                },
                cancellationToken,
                scheduledSlot: ctx).ConfigureAwait(false);
            return;
        }

        var shift = ctx.Shift;
        var date = ctx.ReportDateDdMmmYyyy;

        string? normalizedAttachment =
            !string.IsNullOrEmpty(reportResult.NormalizedZipPath) && File.Exists(reportResult.NormalizedZipPath)
                ? reportResult.NormalizedZipPath
                : (!string.IsNullOrEmpty(reportResult.NormalizedCsvPath) && File.Exists(reportResult.NormalizedCsvPath)
                    ? reportResult.NormalizedCsvPath
                    : null);

        var additional = normalizedAttachment != null
            ? (IReadOnlyList<string>?)new List<string> { normalizedAttachment }
            : null;

        string bodyTemplate = additional != null
            ? _email.CombinedReportBodyWithZip
            : _email.CombinedReportBodyWithoutZip;

        var subject = FormatShiftDateTemplate(_email.CombinedReportSubjectTemplate, shift, date);
        var body = FormatShiftDateTemplate(bodyTemplate, shift, date);

        if (string.IsNullOrWhiteSpace(_email.FromAddress))
        {
            _logger.LogWarning("Email FromAddress is not configured; skipping send.");
            return;
        }

        if (_email.ToAddresses == null || _email.ToAddresses.Count == 0)
        {
            _logger.LogWarning("No ToAddresses configured; skipping send.");
            return;
        }

        var message = new OutgoingEmail
        {
            From = _email.FromAddress,
            To = _email.ToAddresses,
            Cc = _email.CcAddresses is { Count: > 0 } ? _email.CcAddresses : null,
            Subject = subject,
            Body = body,
            PrimaryAttachmentPath = reportResult.ExcelFilePath,
            AdditionalAttachmentPaths = additional
        };

        var sent = await _emailSender.SendAsync(message, cancellationToken).ConfigureAwait(false);
        if (!sent)
        {
            _logger.LogWarning("Combined report step finished but email was not sent successfully.");
            await _missingFileAlerts.SendMissingFilesAlertAsync(new[]
                {
                    $"Scheduled combined report email failed after retries for Shift {shift}, Date {date}."
                },
                cancellationToken,
                scheduledSlot: ctx).ConfigureAwait(false);
        }
        else
            _logger.LogInformation("Combined report and email completed successfully.");
    }

    /// <summary>Gocator merge, then combined Excel and email in one run (used when <c>RunOnStart</c> is true). Gocator-only email is skipped so recipients get one combined mail.</summary>
    public async Task RunFullPipelineAsync(ScheduledJob job, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Running full export pipeline (Gocator merge + combined report + email).");
        var ctx = _scheduleCalculator.ResolveReportContext(job);
        await RunGocatorMergeWithOptionalEmailAsync(job, sendEmailAfterMerge: false, cancellationToken, ctx).ConfigureAwait(false);
        await RunCombinedReportAndEmailAsync(job, cancellationToken, ctx).ConfigureAwait(false);
    }

    public Task RunOnceAsync(CancellationToken cancellationToken = default) =>
        RunFullPipelineAsync(new ScheduledJob(DateTimeOffset.UtcNow, ScheduledJobKind.FullPipeline), cancellationToken);

    private async Task RunGocatorMergeWithOptionalEmailAsync(
        ScheduledJob job,
        bool sendEmailAfterMerge,
        CancellationToken cancellationToken,
        ReportSlotContext? slotContext = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _logger.LogInformation("Running Gocator CSV merge.");
        var ctx = slotContext ?? _scheduleCalculator.ResolveReportContext(job);
        var path = await _gocatorMerge.GenerateCombinedCsvAsync(ctx, cancellationToken).ConfigureAwait(false);

        if (!sendEmailAfterMerge)
            return;

        await TrySendGocatorReportEmailAsync(path, ctx, cancellationToken).ConfigureAwait(false);
    }

    private async Task TrySendGocatorReportEmailAsync(string? csvPath, ReportSlotContext ctx, CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(csvPath) || !File.Exists(csvPath))
        {
            _logger.LogWarning("Gocator CSV was not produced; skipping Gocator email.");
            await _missingFileAlerts.SendMissingFilesAlertAsync(
                new[]
                {
                    $"Gocator scheduled export: no report file for Shift {ctx.Shift}, Date {ctx.ReportDateDdMmmYyyy}."
                },
                cancellationToken,
                scheduledSlot: ctx).ConfigureAwait(false);
            return;
        }

        if (!_email.BypassReportAttachmentSlotCheck &&
            !ReportFileMatchesScheduledSlot(csvPath, ctx))
        {
            _logger.LogWarning(
                "Gocator report file does not match scheduled shift/date; skipping customer email.");
            await _missingFileAlerts.SendMissingFilesAlertAsync(
                new[]
                {
                    $"Gocator report is not for scheduled Shift {ctx.Shift}, Date {ctx.ReportDateDdMmmYyyy} (file: {Path.GetFileName(csvPath)})."
                },
                cancellationToken,
                scheduledSlot: ctx).ConfigureAwait(false);
            return;
        }

        var shift = ctx.Shift;
        var date = ctx.ReportDateDdMmmYyyy;

        if (string.IsNullOrWhiteSpace(_email.FromAddress))
        {
            _logger.LogWarning("Email FromAddress is not configured; skipping Gocator email.");
            return;
        }

        if (_email.ToAddresses == null || _email.ToAddresses.Count == 0)
        {
            _logger.LogWarning("No ToAddresses configured; skipping Gocator email.");
            return;
        }

        var subject = FormatShiftDateTemplate(_email.GocatorReportSubjectTemplate, shift, date);
        var body = string.Format(CultureInfo.InvariantCulture, _email.GocatorReportBodyTemplate, date, shift);

        var message = new OutgoingEmail
        {
            From = _email.FromAddress,
            To = _email.ToAddresses,
            Cc = _email.CcAddresses is { Count: > 0 } ? _email.CcAddresses : null,
            Subject = subject,
            Body = body,
            PrimaryAttachmentPath = csvPath,
            AdditionalAttachmentPaths = null
        };

        var sent = await _emailSender.SendAsync(message, cancellationToken).ConfigureAwait(false);
        if (!sent)
        {
            _logger.LogWarning("Gocator merge finished but Gocator email was not sent successfully.");
            await _missingFileAlerts.SendMissingFilesAlertAsync(new[]
                {
                    $"Scheduled Gocator report email failed after retries for Shift {shift}, Date {date}."
                },
                cancellationToken,
                scheduledSlot: ctx).ConfigureAwait(false);
        }
        else
            _logger.LogInformation("Gocator report email sent successfully.");
    }

    private static bool ReportFileMatchesScheduledSlot(string filePath, ReportSlotContext ctx)
    {
        var name = Path.GetFileNameWithoutExtension(filePath);
        var m = ReportShiftDateInFileName.Match(name);
        if (!m.Success)
            return false;
        if (!ReportCsvDate.ShiftsEquivalent(m.Groups[1].Value.Trim(), ctx.Shift))
            return false;
        if (!ReportCsvDate.TryParseLoose(m.Groups[2].Value.Trim(), out var d))
            return false;
        return d == ctx.ReportDate;
    }

    private static string FormatShiftDateTemplate(string template, string shift, string date)
    {
        if (string.IsNullOrEmpty(template))
            return "";
        return template
            .Replace("{shift}", shift, StringComparison.Ordinal)
            .Replace("{date}", date, StringComparison.Ordinal);
    }
}
