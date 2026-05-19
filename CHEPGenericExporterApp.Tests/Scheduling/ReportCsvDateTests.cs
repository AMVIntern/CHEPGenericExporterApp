using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Tests.Scheduling;

public sealed class ReportCsvDateTests
{
  private static ReportSlotContext Slot(string shift, string dateStr, DateOnly date) =>
      new(shift, dateStr, date);

    [Theory]
    [InlineData("12-MAY-2026", true)]
    [InlineData("12-May-2026", true)]
    [InlineData("2026-05-12", true)]
    [InlineData("", false)]
    [InlineData("not-a-date", false)]
    public void TryParseLoose_parses_common_formats(string input, bool expected)
    {
        var ok = ReportCsvDate.TryParseLoose(input, out var date);
        Assert.Equal(expected, ok);
        if (expected)
            Assert.Equal(new DateOnly(2026, 5, 12), date);
    }

    [Theory]
    [InlineData("3", "3", true)]
    [InlineData("03", "3", true)]
    [InlineData("2", "3", false)]
    public void ShiftsEquivalent_compares_numeric_shifts(string a, string b, bool expected) =>
        Assert.Equal(expected, ReportCsvDate.ShiftsEquivalent(a, b));

    [Fact]
    public void ReportFileMatchesSlot_matches_shift_and_date_in_filename()
    {
        var ctx = Slot("3", "12-MAY-2026", new DateOnly(2026, 5, 12));
        var path = Path.Combine(Path.GetTempPath(), "TEST_Gocator_Report_Shift_3_12-MAY-2026.csv");
        try
        {
            File.WriteAllText(path, "x");
            Assert.True(ReportCsvDate.ReportFileMatchesSlot(path, ctx));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void ReportFileMatchesSlot_rejects_wrong_shift()
    {
        var ctx = Slot("3", "12-MAY-2026", new DateOnly(2026, 5, 12));
        var path = Path.Combine(Path.GetTempPath(), "TEST_Gocator_Report_Shift_2_12-MAY-2026.csv");
        try
        {
            File.WriteAllText(path, "x");
            Assert.False(ReportCsvDate.ReportFileMatchesSlot(path, ctx));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void ReportFileMatchesSlot_rejects_dummy_suffix_in_date_group()
    {
        var ctx = Slot("3", "12-MAY-2026", new DateOnly(2026, 5, 12));
        var path = Path.Combine(Path.GetTempPath(), "Station1_Report_Shift_3_12-MAY-2026_DUMMY.csv");
        try
        {
            File.WriteAllText(path, "x");
            Assert.False(ReportCsvDate.ReportFileMatchesSlot(path, ctx));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }
}
