using CHEPGenericExporterApp.Services.Reports;
using CHEPGenericExporterApp.Services.Scheduling;
using CHEPGenericExporterApp.Tests.Helpers;

namespace CHEPGenericExporterApp.Tests.Reports;

public sealed class StationDummyShiftCsvServiceTests
{
    [Fact]
    public void TryCreateDummyShiftFile_writes_dummy_file_with_template_headers_and_zero_metrics()
    {
        using var env = new ExportTestEnvironment();
        File.Copy(
            TestPaths.Fixture("Stations/Station1_Report_Shift_3_12-May-2026.csv"),
            Path.Combine(env.S1Folder, "Station1_Report_Shift_1_01-Jan-2026.csv"));

        var service = new StationDummyShiftCsvService();
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));
        var gocator = GocatorCsvFactory.CreateMinimal(rowCount: 2);

        var ok = service.TryCreateDummyShiftFile(
            env.S1Folder,
            slot,
            gocator,
            "S1",
            "Station 1 (S1)",
            out var path,
            out var reason);

        Assert.True(ok, reason);
        Assert.NotNull(path);
        Assert.EndsWith("_DUMMY.csv", path, StringComparison.OrdinalIgnoreCase);

        var lines = File.ReadAllLines(path!);
        Assert.True(lines.Length >= 4); // 2 header + 2 data
        Assert.Equal("Date,Timestamp,Shift,Station,RN,PN", lines[0]);
        Assert.All(lines.Skip(2), line =>
        {
            var parts = line.Split(',');
            Assert.Equal("0", parts[4]);
            Assert.Equal("0", parts[5]);
        });
    }

    [Fact]
    public void TryCreateDummyShiftFile_fails_when_no_template_exists()
    {
        using var env = new ExportTestEnvironment();
        var service = new StationDummyShiftCsvService();
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var ok = service.TryCreateDummyShiftFile(
            env.S2Folder,
            slot,
            GocatorCsvFactory.CreateMinimal(),
            "S2",
            "Station 2 (S2)",
            out _,
            out var reason);

        Assert.False(ok);
        Assert.Contains("template", reason, StringComparison.OrdinalIgnoreCase);
    }
}
