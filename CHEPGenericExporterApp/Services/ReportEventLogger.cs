using System.Text.Json;
using System.Text.Json.Serialization;
using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Configuration;

namespace CHEPGenericExporterApp.Services;

public sealed class ReportEventLogger
{
    private readonly string? _logPath;
    private readonly string _siteCode;
    private readonly string _timeZoneId;
    private readonly ILogger<ReportEventLogger> _logger;
    private readonly JsonSerializerOptions _jsonOptions;
    private readonly object _fileLock = new();

    public ReportEventLogger(
        IConfiguration configuration,
        IOptions<ExportPathsOptions> exportPaths,
        IOptions<SchedulerOptions> scheduler,
        ILogger<ReportEventLogger> logger)
    {
        _logPath = configuration["ReportEventLogPath"]?.Trim();
        _siteCode = exportPaths.Value.NormalizedReportSiteCode ?? "";
        _timeZoneId = scheduler.Value.TimeZoneId ?? "Local";
        _logger = logger;
        _jsonOptions = new JsonSerializerOptions
        {
            Converters = { new JsonStringEnumConverter() }
        };
    }

    public void Emit(
        ReportSlotContext ctx,
        ReportEventKind kind,
        ReportOutcomeReason reason,
        string reasonDetail = "",
        IReadOnlyList<string>? missingInputs = null,
        IReadOnlyList<string>? dummyStations = null)
    {
        if (string.IsNullOrWhiteSpace(_logPath))
            return;

        var evt = new ReportEvent
        {
            Timestamp = DateTimeOffset.UtcNow,
            Site = _siteCode,
            Shift = ctx.Shift,
            ReportDate = ctx.ReportDate,
            TimeZone = _timeZoneId,
            Kind = kind,
            Sent = kind is ReportEventKind.GocatorSent or ReportEventKind.CombinedSent,
            Reason = reason,
            ReasonDetail = reasonDetail,
            MissingInputs = missingInputs ?? Array.Empty<string>(),
            DummyStations = dummyStations ?? Array.Empty<string>()
        };

        try
        {
            var dir = Path.GetDirectoryName(_logPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);

            var line = JsonSerializer.Serialize(evt, _jsonOptions);
            lock (_fileLock)
                File.AppendAllText(_logPath, line + Environment.NewLine);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to write report event log entry.");
        }
    }
}
