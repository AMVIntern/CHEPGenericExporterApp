using System.Globalization;
using System.Text.RegularExpressions;

namespace CHEPGenericExporterApp.Services.Scheduling;

/// <summary>Shift and calendar date for the scheduled export slot (Sydney wall date).</summary>
public readonly record struct ReportSlotContext(string Shift, string ReportDateDdMmmYyyy, DateOnly ReportDate);

/// <summary>Parses and compares dates as they appear in CSV cells and report filenames.</summary>
public static class ReportCsvDate
{
    private static readonly Regex ReportShiftDateInFileName = new(
        @"_Shift_(\d+)_(.+)$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>True when the report file name ends with <c>_Shift_{n}_{date}</c> matching <paramref name="ctx"/>.</summary>
    public static bool ReportFileMatchesSlot(string filePath, ReportSlotContext ctx)
    {
        var name = Path.GetFileNameWithoutExtension(filePath);
        var m = ReportShiftDateInFileName.Match(name);
        if (!m.Success)
            return false;
        if (!ShiftsEquivalent(m.Groups[1].Value.Trim(), ctx.Shift))
            return false;
        if (!TryParseLoose(m.Groups[2].Value.Trim(), out var d))
            return false;
        return d == ctx.ReportDate;
    }

    public static bool TryParseLoose(string? s, out DateOnly date)
    {
        date = default;
        if (string.IsNullOrWhiteSpace(s))
            return false;

        s = s.Trim();
        if (DateTime.TryParseExact(s, "dd-MMM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
        {
            date = DateOnly.FromDateTime(dt);
            return true;
        }

        if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
        {
            date = DateOnly.FromDateTime(dt);
            return true;
        }

        return false;
    }

    public static bool EqualsLoose(string? a, string? b)
    {
        if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
            return false;
        if (string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase))
            return true;
        return TryParseLoose(a, out var da) && TryParseLoose(b, out var db) && da == db;
    }

    public static bool MatchesContext(string? cellOrFragment, ReportSlotContext ctx) =>
        TryParseLoose(cellOrFragment, out var d) && d == ctx.ReportDate;

    public static bool ShiftsEquivalent(string? dataShift, string expectedShift)
    {
        if (string.IsNullOrWhiteSpace(dataShift) || string.IsNullOrWhiteSpace(expectedShift))
            return false;
        if (int.TryParse(dataShift.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var a) &&
            int.TryParse(expectedShift.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var b))
            return a == b;
        return string.Equals(dataShift.Trim(), expectedShift.Trim(), StringComparison.OrdinalIgnoreCase);
    }
}
