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
    public void GetNextScheduledJob_skips_saturday()
    {
        var calc = CreateCalculator();
        // 2026-05-16 is Saturday UTC
        var after = new DateTimeOffset(2026, 5, 16, 1, 0, 0, TimeSpan.Zero);
        var next = calc.GetNextScheduledJob(after);

        Assert.True(next.Utc > after);
        var local = TimeZoneInfo.ConvertTimeFromUtc(next.Utc.UtcDateTime, TimeZoneInfo.Utc);
        Assert.NotEqual(DayOfWeek.Saturday, local.DayOfWeek);
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
}
