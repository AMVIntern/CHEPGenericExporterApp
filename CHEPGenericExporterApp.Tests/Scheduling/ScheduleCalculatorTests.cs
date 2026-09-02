using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Tests.Scheduling;

public sealed class ScheduleCalculatorTests
{
    private static ScheduleCalculator CreateCalculator() =>
        new(Options.Create(new SchedulerOptions
        {
            TimeZoneId = "UTC",
            GocatorTimes = new List<string> { "06:00", "14:00", "22:00" },
            CombinedTimes = new List<string> { "06:02", "14:02", "22:02" }
        }));

    [Fact]
    public void Constructor_throws_when_gocator_and_combined_counts_differ()
    {
        Assert.Throws<InvalidOperationException>(() => _ = new ScheduleCalculator(Options.Create(new SchedulerOptions
        {
            TimeZoneId = "UTC",
            GocatorTimes = new List<string> { "06:00" },
            CombinedTimes = new List<string> { "06:02", "14:02" }
        })));
    }

    [Fact]
    public void ResolveReportContext_maps_first_slot_to_shift3_previous_day()
    {
        var calc = CreateCalculator();
        var job = new ScheduledJob(
            new DateTimeOffset(2026, 5, 13, 6, 1, 0, TimeSpan.Zero),
            ScheduledJobKind.CombinedReportAndEmail);

        var ctx = calc.ResolveReportContext(job);

        Assert.Equal("3", ctx.Shift);
        Assert.Equal(new DateOnly(2026, 5, 12), ctx.ReportDate);
    }

    [Fact]
    public void ResolveReportContext_maps_second_slot_to_shift1_same_day()
    {
        var calc = CreateCalculator();
        var job = new ScheduledJob(
            new DateTimeOffset(2026, 5, 13, 14, 1, 0, TimeSpan.Zero),
            ScheduledJobKind.GocatorMerge);

        var ctx = calc.ResolveReportContext(job);

        Assert.Equal("1", ctx.Shift);
        Assert.Equal(new DateOnly(2026, 5, 13), ctx.ReportDate);
    }

    [Fact]
    public void GetNextScheduledJob_returns_gocator_before_combined_on_weekday()
    {
        var calc = CreateCalculator();
        var after = new DateTimeOffset(2026, 5, 12, 5, 0, 0, TimeSpan.Zero); // Tuesday
        var first = calc.GetNextScheduledJob(after);
        Assert.Equal(ScheduledJobKind.GocatorMerge, first.Kind);

        var second = calc.GetNextScheduledJob(first.Utc);
        Assert.Equal(ScheduledJobKind.CombinedReportAndEmail, second.Kind);
        Assert.True(second.Utc > first.Utc);
    }

    /// <summary>
    /// Walks a full Mon–Sun week (2026-05-11 Mon .. 2026-05-17 Sun, UTC) and collects every
    /// (Shift, ReportDate) pair the scheduler produces via <see cref="ScheduleCalculator.GetNextScheduledJob"/>
    /// + <see cref="ScheduleCalculator.ResolveReportContext"/> — i.e. it exercises the exact same two calls
    /// production code uses, not a re-implementation of the scheduling rules. Asserts the complete set,
    /// not just presence/absence of individual days, so a regression that shifts a report to the wrong
    /// date (e.g. "Saturday's Shift 3" landing on Friday instead of Saturday) fails loudly.
    /// </summary>
    [Fact]
    public void GetNextScheduledJob_produces_the_expected_shift_set_for_a_full_week()
    {
        var calc = CreateCalculator();
        var cursor = new DateTimeOffset(2026, 5, 11, 0, 0, 0, TimeSpan.Zero); // Monday 00:00 UTC
        var weekEnd = cursor.AddDays(7);

        var combinedJobs = new List<(DayOfWeek FiredOn, string Shift, DateOnly ReportDate)>();

        while (cursor < weekEnd)
        {
            var job = calc.GetNextScheduledJob(cursor);
            if (job.Utc >= weekEnd)
                break;

            if (job.Kind == ScheduledJobKind.CombinedReportAndEmail)
            {
                var ctx = calc.ResolveReportContext(job);
                var firedOn = TimeZoneInfo.ConvertTimeFromUtc(job.Utc.UtcDateTime, TimeZoneInfo.Utc).DayOfWeek;
                combinedJobs.Add((firedOn, ctx.Shift, ctx.ReportDate));
            }

            cursor = job.Utc;
        }

        // Monday–Friday: unchanged, every day gets Shift 3 (of the previous day), Shift 1, Shift 2.
        foreach (var date in new[]
                 {
                     new DateOnly(2026, 5, 11), // Monday
                     new DateOnly(2026, 5, 12), // Tuesday
                     new DateOnly(2026, 5, 13), // Wednesday
                     new DateOnly(2026, 5, 14), // Thursday
                     new DateOnly(2026, 5, 15), // Friday
                 })
        {
            Assert.Contains(combinedJobs, e => e.Shift == "1" && e.ReportDate == date);
            Assert.Contains(combinedJobs, e => e.Shift == "2" && e.ReportDate == date);
        }

        // Sunday's own Shift 3 (2026-05-17) is produced only by the *following* Monday's slot 0 —
        // outside this week's window — so it must not appear here. This is the pre-existing,
        // untouched behaviour the requirement calls "already handled by the existing system".
        Assert.DoesNotContain(combinedJobs, e => e.Shift == "3" && e.ReportDate == new DateOnly(2026, 5, 17));

        // Friday's Shift 3 (2026-05-15) must NEVER appear. It would be produced by a slot firing at
        // 06:00 Saturday — and that slot must stay excluded, exactly as it always has been (Saturday
        // was fully skipped before this feature; it must still not produce this specific report now).
        Assert.DoesNotContain(combinedJobs, e => e.Shift == "3" && e.ReportDate == new DateOnly(2026, 5, 15));

        // Saturday (2026-05-16): Shift 1 and Shift 2 fire on Saturday itself.
        Assert.Contains(combinedJobs, e => e.FiredOn == DayOfWeek.Saturday && e.Shift == "1" && e.ReportDate == new DateOnly(2026, 5, 16));
        Assert.Contains(combinedJobs, e => e.FiredOn == DayOfWeek.Saturday && e.Shift == "2" && e.ReportDate == new DateOnly(2026, 5, 16));

        // Saturday's Shift 3 fires as part of Sunday's slot 0 (06:00 Sunday), dated Saturday —
        // this is what actually satisfies "enable Saturday Shift 3".
        Assert.Contains(combinedJobs, e => e.FiredOn == DayOfWeek.Sunday && e.Shift == "3" && e.ReportDate == new DateOnly(2026, 5, 16));

        // Sunday (2026-05-17): Shift 1 and Shift 2 fire on Sunday itself.
        Assert.Contains(combinedJobs, e => e.FiredOn == DayOfWeek.Sunday && e.Shift == "1" && e.ReportDate == new DateOnly(2026, 5, 17));
        Assert.Contains(combinedJobs, e => e.FiredOn == DayOfWeek.Sunday && e.Shift == "2" && e.ReportDate == new DateOnly(2026, 5, 17));

        // Exactly 20 combined-report jobs in the week: 5 weekdays x 3 (Shift 3-of-previous-day,
        // Shift 1, Shift 2) + Saturday's own 2 (Shift 1, 2) + Sunday's 3 (Shift 3-of-Saturday,
        // Shift 1, Shift 2) = 15 + 2 + 3 = 20. No duplicates.
        Assert.Equal(20, combinedJobs.Count);
        Assert.Equal(combinedJobs.Count, combinedJobs.Distinct().Count());
    }
}
