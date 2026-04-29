using System.Globalization;
using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Services;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.HostedServices;

/// <summary>Every configured interval, retries unsent shift emails that are already overdue.</summary>
public sealed class MissedEmailRecoveryWorker : BackgroundService
{
    private readonly CsvAuditLogger _csvAuditLogger;
    private readonly ExportPipeline _pipeline;
    private readonly SchedulerOptions _schedulerOptions;
    private readonly int _intervalMinutes;
    private readonly ILogger<MissedEmailRecoveryWorker> _logger;

    public MissedEmailRecoveryWorker(
        CsvAuditLogger csvAuditLogger,
        ExportPipeline pipeline,
        IOptions<SchedulerOptions> schedulerOptions,
        IConfiguration configuration,
        ILogger<MissedEmailRecoveryWorker> logger)
    {
        _csvAuditLogger = csvAuditLogger;
        _pipeline = pipeline;
        _schedulerOptions = schedulerOptions.Value;
        _intervalMinutes = _csvAuditLogger.GetMissedSendCheckIntervalMinutes(configuration);
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                await RecoverMissedSendsAsync(stoppingToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Missed-send recovery cycle failed.");
            }

            await Task.Delay(TimeSpan.FromMinutes(_intervalMinutes), stoppingToken).ConfigureAwait(false);
        }
    }

    private async Task RecoverMissedSendsAsync(CancellationToken cancellationToken)
    {
        var pendingRows = _csvAuditLogger.GetPendingRows();
        if (pendingRows.Count == 0)
            return;

        var gocatorTimes = ParseShiftTimes(_schedulerOptions.GocatorTimes, "Scheduler:GocatorTimes");
        var combinedTimes = ParseShiftTimes(_schedulerOptions.CombinedTimes, "Scheduler:CombinedTimes");

        foreach (var row in pendingRows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var shiftSlotIndex = GetSlotIndexForShift(row.Shift);
            var nowLocal = DateTime.Now;
            var context = new ReportSlotContext(
                row.Shift,
                row.Date.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture),
                row.Date);

            if (!row.GocatorReportSent)
            {
                var gocatorScheduled = row.Date.ToDateTime(TimeOnly.FromTimeSpan(gocatorTimes[shiftSlotIndex]));
                if (nowLocal > gocatorScheduled)
                {
                    _csvAuditLogger.MarkAttempt(context.Shift, context.ReportDate);
                    var targetDate = row.Date.ToDateTime(TimeOnly.MinValue);
                    var sent = await _pipeline.TryRecoverGocatorEmailAsync(context, cancellationToken, targetDate).ConfigureAwait(false);
                    if (sent)
                    {
                        _logger.LogInformation("Recovered Gocator email for Shift {Shift}, Date {Date}.", context.Shift, context.ReportDateDdMmmYyyy);
                    }
                }
            }

            if (!row.CombinedReportSent)
            {
                var combinedScheduled = row.Date.ToDateTime(TimeOnly.FromTimeSpan(combinedTimes[shiftSlotIndex]));
                if (nowLocal > combinedScheduled)
                {
                    _csvAuditLogger.MarkAttempt(context.Shift, context.ReportDate);
                    var sent = await _pipeline.TryRecoverCombinedEmailAsync(context, cancellationToken).ConfigureAwait(false);
                    if (sent)
                    {
                        _logger.LogInformation("Recovered combined email for Shift {Shift}, Date {Date}.", context.Shift, context.ReportDateDdMmmYyyy);
                    }
                }
            }
        }
    }

    private static TimeSpan[] ParseShiftTimes(List<string>? configured, string settingName)
    {
        if (configured is null || configured.Count != 3)
            throw new InvalidOperationException($"{settingName} must contain exactly 3 entries (Shift 1..3).");

        var result = new TimeSpan[3];
        for (var i = 0; i < configured.Count; i++)
        {
            if (!TimeSpan.TryParse(configured[i], CultureInfo.InvariantCulture, out var parsed))
                throw new InvalidOperationException($"{settingName}[{i}] value '{configured[i]}' is not a valid time.");
            result[i] = parsed;
        }

        return result;
    }

    private static int GetSlotIndexForShift(string shiftRaw)
    {
        if (!int.TryParse(shiftRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var shift))
            throw new InvalidOperationException($"Invalid shift value '{shiftRaw}' in audit row.");

        return shift switch
        {
            // Align with ScheduleCalculator 3-slot CHEP mapping:
            // slot 0 (06:xx) -> Shift 3
            // slot 1 (14:xx) -> Shift 1
            // slot 2 (22:xx) -> Shift 2
            3 => 0,
            1 => 1,
            2 => 2,
            _ => throw new InvalidOperationException($"Unsupported shift '{shiftRaw}' in audit row.")
        };
    }
}
