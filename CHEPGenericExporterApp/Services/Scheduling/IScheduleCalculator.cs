namespace CHEPGenericExporterApp.Services.Scheduling;

/// <summary>Computes the next scheduled job after a reference UTC time.</summary>
public interface IScheduleCalculator
{
    /// <summary>Next job time (UTC) strictly after <paramref name="afterUtc"/>, and what to run.</summary>
    ScheduledJob GetNextScheduledJob(DateTimeOffset afterUtc);

    /// <returns>UTC instant strictly after <paramref name="afterUtc"/>.</returns>
    DateTimeOffset GetNextRunUtc(DateTimeOffset afterUtc) => GetNextScheduledJob(afterUtc).Utc;
}
