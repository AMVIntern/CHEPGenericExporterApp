using System.Globalization;
using CHEPGenericExporterApp.Configuration;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Services.Scheduling;

public sealed class ScheduleCalculator : IScheduleCalculator
{
    private static readonly TimeSpan[] DefaultGocatorTimes = [new(6, 0, 0), new(14, 0, 0), new(22, 0, 0)];
    private static readonly TimeSpan[] DefaultCombinedTimes = [new(6, 2, 0), new(14, 2, 0), new(22, 2, 0)];

    private readonly SchedulerOptions _options;
    private readonly TimeZoneInfo _timeZone;
    private readonly TimeSpan[] _gocatorSlots;
    private readonly TimeSpan[] _combinedSlots;

    public ScheduleCalculator(IOptions<SchedulerOptions> options)
    {
        _options = options.Value ?? throw new ArgumentNullException(nameof(options));
        _timeZone = ResolveTimeZone(_options.TimeZoneId);

        _gocatorSlots = ParseTimeList(_options.GocatorTimes, DefaultGocatorTimes);
        _combinedSlots = ParseTimeList(_options.CombinedTimes, DefaultCombinedTimes);

        if (_gocatorSlots.Length != _combinedSlots.Length)
            throw new InvalidOperationException("Scheduler: GocatorTimes and CombinedTimes must have the same number of entries.");
        if (_gocatorSlots.Length == 0)
            throw new InvalidOperationException("Scheduler: at least one time slot pair is required.");
    }

    public ScheduledJob GetNextScheduledJob(DateTimeOffset afterUtc) => GetNextAmvJob(afterUtc);

    public ReportSlotContext ResolveReportContext(ScheduledJob job)
    {
        var local = TimeZoneInfo.ConvertTimeFromUtc(job.Utc.UtcDateTime, _timeZone);
        var dateOnly = DateOnly.FromDateTime(local.Date);

        var slots = job.Kind switch
        {
            ScheduledJobKind.CombinedReportAndEmail => _combinedSlots,
            ScheduledJobKind.FullPipeline => _gocatorSlots,
            _ => _gocatorSlots,
        };

        var idx = FindMatchingSlotIndex(local, slots, toleranceMinutes: 3);
        if (idx < 0)
            idx = FindClosestSlotIndex(local, slots);

        // CHEP three-run day (06:00, 14:00, 22:00 in config order): see MapChepThreeSlotToShiftAndDate.
        var (shift, reportDate) = MapChepThreeSlotToShiftAndDate(idx, dateOnly);

        var dateStr = reportDate.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);
        return new ReportSlotContext(shift, dateStr, reportDate);
    }

    /// <summary>
    /// When exactly three schedule pairs exist (typical 6 / 14 / 22): first run → Shift 3, previous day;
    /// second → Shift 1, same day; third → Shift 2, same day. Otherwise falls back to slot index + 1 and job date.
    /// </summary>
    private (string Shift, DateOnly ReportDate) MapChepThreeSlotToShiftAndDate(int slotIndex, DateOnly jobLocalDate)
    {
        if (_gocatorSlots.Length != 3)
            return ((slotIndex + 1).ToString(CultureInfo.InvariantCulture), jobLocalDate);

        return slotIndex switch
        {
            0 => ("3", jobLocalDate.AddDays(-1)),
            1 => ("1", jobLocalDate),
            2 => ("2", jobLocalDate),
            _ => ((slotIndex + 1).ToString(CultureInfo.InvariantCulture), jobLocalDate),
        };
    }

    private static int FindMatchingSlotIndex(DateTime local, TimeSpan[] slots, double toleranceMinutes)
    {
        var tod = local.TimeOfDay;
        for (var i = 0; i < slots.Length; i++)
        {
            var diff = Math.Abs((tod - slots[i]).TotalMinutes);
            if (diff <= toleranceMinutes)
                return i;
        }

        return -1;
    }

    private static int FindClosestSlotIndex(DateTime local, TimeSpan[] slots)
    {
        var tod = local.TimeOfDay;
        var best = 0;
        var bestDiff = double.MaxValue;
        for (var i = 0; i < slots.Length; i++)
        {
            var diff = Math.Abs((tod - slots[i]).TotalMinutes);
            if (diff < bestDiff)
            {
                bestDiff = diff;
                best = i;
            }
        }

        return best;
    }

    /// <summary>Saturday skipped. Sunday: last slot pair only. Mon–Fri: all pairs (Gocator then combined).</summary>
    private ScheduledJob GetNextAmvJob(DateTimeOffset afterUtc)
    {
        var now = TimeZoneInfo.ConvertTime(afterUtc, _timeZone).DateTime;

        for (var guard = 0; guard < 400; guard++)
        {
            var day = now.DayOfWeek;
            var date = now.Date;

            if (day == DayOfWeek.Saturday)
            {
                now = date.AddDays(1);
                continue;
            }

            var events = new List<(DateTimeOffset Utc, ScheduledJobKind Kind)>(6);
            AppendAmvEventsForDay(date, day, events);
            events.Sort((a, b) => a.Utc.CompareTo(b.Utc));

            foreach (var (utc, kind) in events)
            {
                if (utc > afterUtc)
                    return new ScheduledJob(utc, kind);
            }

            now = date.AddDays(1);
        }

        throw new InvalidOperationException("Scheduler: could not compute next run.");
    }

    private void AppendAmvEventsForDay(DateTime date, DayOfWeek day, List<(DateTimeOffset Utc, ScheduledJobKind Kind)> events)
    {
        if (day == DayOfWeek.Saturday)
            return;

        var g = _gocatorSlots;
        var c = _combinedSlots;

        if (day == DayOfWeek.Sunday)
        {
            var i = g.Length - 1;
            events.Add((WallTimeInZoneToUtc(date + g[i]), ScheduledJobKind.GocatorMerge));
            events.Add((WallTimeInZoneToUtc(date + c[i]), ScheduledJobKind.CombinedReportAndEmail));
            return;
        }

        if (day is >= DayOfWeek.Monday and <= DayOfWeek.Friday)
        {
            for (var i = 0; i < g.Length; i++)
            {
                events.Add((WallTimeInZoneToUtc(date + g[i]), ScheduledJobKind.GocatorMerge));
                events.Add((WallTimeInZoneToUtc(date + c[i]), ScheduledJobKind.CombinedReportAndEmail));
            }
        }
    }

    private DateTimeOffset WallTimeInZoneToUtc(DateTime localDateTimeUnspecified)
    {
        var unspecified = DateTime.SpecifyKind(localDateTimeUnspecified, DateTimeKind.Unspecified);
        var utc = TimeZoneInfo.ConvertTimeToUtc(unspecified, _timeZone);
        return new DateTimeOffset(utc, TimeSpan.Zero);
    }

    private static TimeZoneInfo ResolveTimeZone(string? timeZoneId)
    {
        var id = (timeZoneId ?? "Local").Trim();
        if (id.Equals("Local", StringComparison.OrdinalIgnoreCase))
            return TimeZoneInfo.Local;

        try
        {
            return TimeZoneInfo.FindSystemTimeZoneById(id);
        }
        catch (TimeZoneNotFoundException)
        {
            throw new InvalidOperationException(
                $"Scheduler: unknown TimeZoneId '{id}'. Use an IANA ID on Linux/macOS or a Windows ID on Windows, or 'Local'.");
        }
    }

    private static TimeSpan[] ParseTimeList(List<string>? list, TimeSpan[] defaults)
    {
        if (list == null || list.Count == 0)
            return defaults;
        return list.Select(ParseTimeOfDay).ToArray();
    }

    private static TimeSpan ParseTimeOfDay(string raw)
    {
        var s = raw?.Trim() ?? "";
        if (TimeSpan.TryParse(s, out var ts))
            return ts;

        var parts = s.Split(':', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 2 &&
            int.TryParse(parts[0], out var h) &&
            int.TryParse(parts[1], out var m))
        {
            var sec = parts.Length > 2 && int.TryParse(parts[2], out var s0) ? s0 : 0;
            return new TimeSpan(h, m, sec);
        }

        throw new InvalidOperationException($"Scheduler: invalid time entry '{raw}'. Use HH:mm or HH:mm:ss.");
    }
}
