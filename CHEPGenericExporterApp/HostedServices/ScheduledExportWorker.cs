using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Services;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.HostedServices;

/// <summary>Waits for the next schedule instant, then runs the appropriate export step until shutdown.</summary>
public sealed class ScheduledExportWorker : BackgroundService
{
    private readonly IScheduleCalculator _scheduleCalculator;
    private readonly ExportPipeline _pipeline;
    private readonly CsvAuditLogger _csvAuditLogger;
    private readonly IOptions<SchedulerOptions> _schedulerOptions;
    private readonly ILogger<ScheduledExportWorker> _logger;

    public ScheduledExportWorker(
        IScheduleCalculator scheduleCalculator,
        ExportPipeline pipeline,
        CsvAuditLogger csvAuditLogger,
        IOptions<SchedulerOptions> schedulerOptions,
        ILogger<ScheduledExportWorker> logger)
    {
        _scheduleCalculator = scheduleCalculator;
        _pipeline = pipeline;
        _csvAuditLogger = csvAuditLogger;
        _schedulerOptions = schedulerOptions;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var opts = _schedulerOptions.Value;
        var useImmediateFirstRun = opts.RunOnStart;
        EnsureAuditRowForNextUpcomingShift();

        while (!stoppingToken.IsCancellationRequested)
        {
            ScheduledJob nextJob;
            if (useImmediateFirstRun)
            {
                nextJob = new ScheduledJob(DateTimeOffset.UtcNow, ScheduledJobKind.FullPipeline);
                useImmediateFirstRun = false;
                _logger.LogInformation("RunOnStart enabled: executing full pipeline immediately.");
            }
            else
            {
                var after = DateTimeOffset.UtcNow;
                nextJob = _scheduleCalculator.GetNextScheduledJob(after);
                EnsureAuditRow(nextJob);
                var delay = nextJob.Utc - DateTimeOffset.UtcNow;
                if (delay > TimeSpan.Zero)
                {
                    var tz = ResolveSchedulerTimeZone();
                    var nextLocal = TimeZoneInfo.ConvertTimeFromUtc(nextJob.Utc.UtcDateTime, tz);
                    _logger.LogInformation(
                        "Next job {Job} at {LocalTime:yyyy-MM-dd HH:mm:ss} ({TimeZoneId}, local). Waiting {Minutes:F1} minutes.",
                        nextJob.Kind,
                        nextLocal,
                        tz.Id,
                        delay.TotalMinutes);
                    await WaitWithPeriodicLogsAsync(delay, tz, stoppingToken).ConfigureAwait(false);
                }
            }

            try
            {
                await RunJobAsync(nextJob, stoppingToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unhandled error during scheduled job {Job}.", nextJob.Kind);
            }
        }
    }

    private void EnsureAuditRowForNextUpcomingShift()
    {
        try
        {
            var next = _scheduleCalculator.GetNextScheduledJob(DateTimeOffset.UtcNow);
            EnsureAuditRow(next);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not ensure startup audit row for next scheduled shift.");
        }
    }

    private void EnsureAuditRow(ScheduledJob job)
    {
        var context = _scheduleCalculator.ResolveReportContext(job);
        _csvAuditLogger.EnsureRow(context.Shift, context.ReportDate);
    }

    private async Task RunJobAsync(ScheduledJob job, CancellationToken stoppingToken)
    {
        switch (job.Kind)
        {
            case ScheduledJobKind.GocatorMerge:
                await _pipeline.RunGocatorMergeOnlyAsync(job, stoppingToken).ConfigureAwait(false);
                break;
            case ScheduledJobKind.CombinedReportAndEmail:
                await _pipeline.RunCombinedReportAndEmailAsync(job, stoppingToken).ConfigureAwait(false);
                break;
            case ScheduledJobKind.FullPipeline:
                await _pipeline.RunFullPipelineAsync(job, stoppingToken).ConfigureAwait(false);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(job), job.Kind, null);
        }
    }

    private async Task WaitWithPeriodicLogsAsync(TimeSpan totalDelay, TimeZoneInfo displayTimeZone, CancellationToken stoppingToken)
    {
        var opts = _schedulerOptions.Value;
        var chunk = TimeSpan.FromSeconds(Math.Clamp(opts.StatusLogIntervalSeconds, 5, 3600));
        var remaining = totalDelay;

        while (remaining > TimeSpan.Zero && !stoppingToken.IsCancellationRequested)
        {
            var step = remaining > chunk ? chunk : remaining;
            try
            {
                await Task.Delay(step, stoppingToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }

            remaining -= step;
            if (remaining > TimeSpan.Zero)
            {
                var nowLocal = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, displayTimeZone);
                _logger.LogInformation(
                    "Remaining wait: {Minutes:F1} minutes (at {LocalTime:yyyy-MM-dd HH:mm:ss} {TimeZoneId}).",
                    remaining.TotalMinutes,
                    nowLocal,
                    displayTimeZone.Id);
            }
        }
    }

    private TimeZoneInfo ResolveSchedulerTimeZone()
    {
        var id = _schedulerOptions.Value.TimeZoneId?.Trim();
        if (string.IsNullOrEmpty(id) || id.Equals("Local", StringComparison.OrdinalIgnoreCase))
            return TimeZoneInfo.Local;

        try
        {
            return TimeZoneInfo.FindSystemTimeZoneById(id);
        }
        catch (TimeZoneNotFoundException)
        {
            _logger.LogWarning("Scheduler TimeZoneId '{TimeZoneId}' not found; using local time for logs.", id);
            return TimeZoneInfo.Local;
        }
    }
}
