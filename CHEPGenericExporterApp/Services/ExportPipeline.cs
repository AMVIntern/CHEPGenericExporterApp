using System.Globalization;
using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Reports;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Services;

/// <summary>Runs Gocator CSV merge, combined Excel generation, and optional email send.</summary>
public sealed class ExportPipeline
{
    private readonly GocatorCsvMergeService _gocatorMerge;
    private readonly CombinedExcelReportService _excelReport;
    private readonly IEmailSender _emailSender;
    private readonly EmailOptions _email;
    private readonly ILogger<ExportPipeline> _logger;

    public ExportPipeline(
        GocatorCsvMergeService gocatorMerge,
        CombinedExcelReportService excelReport,
        IEmailSender emailSender,
        IOptions<EmailOptions> emailOptions,
        ILogger<ExportPipeline> logger)
    {
        _gocatorMerge = gocatorMerge;
        _excelReport = excelReport;
        _emailSender = emailSender;
        _email = emailOptions.Value;
        _logger = logger;
    }

    /// <summary>Gocator merge for the scheduled slot, then email with the merged CSV (same idea as combined + email).</summary>
    public Task RunGocatorMergeOnlyAsync(CancellationToken cancellationToken = default) =>
        RunGocatorMergeWithOptionalEmailAsync(sendEmailAfterMerge: true, cancellationToken);

    /// <summary>Combined Excel from latest Gocator CSV, then email (second step of each shift window).</summary>
    public async Task RunCombinedReportAndEmailAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Running combined Excel report and email.");

        var reportResult = _excelReport.GenerateCombinedExcelReport();

        if (reportResult == null ||
            string.IsNullOrEmpty(reportResult.ExcelFilePath) ||
            !File.Exists(reportResult.ExcelFilePath))
        {
            _logger.LogWarning("Combined Excel report was not produced; skipping email.");
            return;
        }

        var (shift, date) = ParseShiftAndDateFromShiftFileName(reportResult.ExcelFilePath);

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
            _logger.LogWarning("Combined report step finished but email was not sent successfully.");
        else
            _logger.LogInformation("Combined report and email completed successfully.");
    }

    /// <summary>Gocator merge, then combined Excel and email in one run (used when <c>RunOnStart</c> is true). Gocator-only email is skipped so recipients get one combined mail.</summary>
    public async Task RunFullPipelineAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Running full export pipeline (Gocator merge + combined report + email).");
        await RunGocatorMergeWithOptionalEmailAsync(sendEmailAfterMerge: false, cancellationToken).ConfigureAwait(false);
        await RunCombinedReportAndEmailAsync(cancellationToken).ConfigureAwait(false);
    }

    public Task RunOnceAsync(CancellationToken cancellationToken = default) =>
        RunFullPipelineAsync(cancellationToken);

    private async Task RunGocatorMergeWithOptionalEmailAsync(bool sendEmailAfterMerge, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _logger.LogInformation("Running Gocator CSV merge.");
        var path = _gocatorMerge.GenerateCombinedCsv();

        if (!sendEmailAfterMerge)
            return;

        await TrySendGocatorReportEmailAsync(path, cancellationToken).ConfigureAwait(false);
    }

    private async Task TrySendGocatorReportEmailAsync(string? csvPath, CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(csvPath) || !File.Exists(csvPath))
        {
            _logger.LogWarning("Gocator CSV was not produced; skipping Gocator email.");
            return;
        }

        var (shift, date) = ParseShiftAndDateFromShiftFileName(csvPath);

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
            _logger.LogWarning("Gocator merge finished but Gocator email was not sent successfully.");
        else
            _logger.LogInformation("Gocator report email sent successfully.");
    }

    /// <summary>Parses <c>...Shift_{shift}_{date}</c> from Gocator CSV or combined Excel file names.</summary>
    private static (string shift, string date) ParseShiftAndDateFromShiftFileName(string filePath)
    {
        string fileName = Path.GetFileNameWithoutExtension(filePath);
        string shift = "Unknown";
        string date = DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);

        if (fileName.Contains("Shift_", StringComparison.Ordinal))
        {
            int shiftIndex = fileName.IndexOf("Shift_", StringComparison.Ordinal) + 6;
            int underscoreIndex = fileName.IndexOf('_', shiftIndex);
            if (underscoreIndex > shiftIndex)
                shift = fileName.Substring(shiftIndex, underscoreIndex - shiftIndex);

            int lastUnderscore = fileName.LastIndexOf('_');
            if (lastUnderscore > 0)
                date = fileName[(lastUnderscore + 1)..];
        }

        return (shift, date);
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
