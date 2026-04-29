namespace CHEPGenericExporterApp.Configuration;

/// <summary>AMV shift schedule: Saturday skipped; Sunday uses the last time pair only; Monday–Friday runs all pairs (Gocator merge, then combined report + email).</summary>
public sealed class SchedulerOptions
{
    public const string SectionName = "Scheduler";

    /// <summary>IANA ID (e.g. Australia/Sydney) or "Local".</summary>
    public string TimeZoneId { get; set; } = "Local";

    /// <summary>
    /// Wall times for Gocator CSV merge (default 06:00, 14:00, 22:00).
    /// Index 0 maps to Shift 1, index 1 to Shift 2, index 2 to Shift 3.
    /// </summary>
    public List<string>? GocatorTimes { get; set; }

    /// <summary>Wall times for combined Excel + email (default 06:02, 14:02, 22:02). Must match <see cref="GocatorTimes"/> count.</summary>
    public List<string>? CombinedTimes { get; set; }

    /// <summary>If true, runs the full pipeline once immediately on startup, then follows the schedule.</summary>
    public bool RunOnStart { get; set; }

    /// <summary>How often to log remaining wait time while delaying until the next run.</summary>
    public int StatusLogIntervalSeconds { get; set; } = 30;
}
