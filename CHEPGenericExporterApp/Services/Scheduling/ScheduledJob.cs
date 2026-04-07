namespace CHEPGenericExporterApp.Services.Scheduling;

public enum ScheduledJobKind
{
    /// <summary>Gocator Top/Bottom CSV merge only.</summary>
    GocatorMerge,

    /// <summary>Combined Excel + email (after Gocator merge in the same shift window).</summary>
    CombinedReportAndEmail,

    /// <summary>Gocator merge, then combined Excel and email (used for RunOnStart).</summary>
    FullPipeline
}

public readonly record struct ScheduledJob(DateTimeOffset Utc, ScheduledJobKind Kind);
