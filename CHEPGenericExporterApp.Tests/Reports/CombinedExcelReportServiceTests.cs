using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Reports;
using CHEPGenericExporterApp.Services.Scheduling;
using CHEPGenericExporterApp.Tests.Helpers;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Moq;

namespace CHEPGenericExporterApp.Tests.Reports;

public sealed class CombinedExcelReportServiceTests
{
    [Fact]
    public async Task GenerateCombinedExcelReportAsync_builds_excel_and_power_bi_csv_when_all_inputs_present()
    {
        using var env = await TestDataSeeder.SeedFullEnvironmentAsync();
        var service = CreateService(env);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedExcelReportAsync(slot);

        Assert.NotNull(result);
        Assert.NotEmpty(result!.ExcelFilePath);
        Assert.True(File.Exists(result.ExcelFilePath));
        Assert.NotEmpty(result.NormalizedCsvPath);
        Assert.True(File.Exists(result.NormalizedCsvPath));
        Assert.Empty(result.DummyStationsUsed);
        Assert.Contains("TEST_Combined_Report_Shift_3_12-MAY-2026.xlsx", result.ExcelFilePath);
    }

    [Fact]
    public async Task GenerateCombinedExcelReportAsync_returns_no_excel_path_with_missing_inputs_when_gocator_missing()
    {
        using var env = new ExportTestEnvironment();
        TestDataSeeder.CopyStationFiles(env);
        var service = CreateService(env);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedExcelReportAsync(slot);

        Assert.NotNull(result);
        Assert.Null(result!.ExcelFilePath);
        Assert.NotEmpty(result.MissingInputs);
    }

    [Fact]
    public async Task GenerateCombinedExcelReportAsync_uses_dummy_station_file_when_enabled()
    {
        using var env = await TestDataSeeder.SeedFullEnvironmentAsync(includeS2: false);
        File.Copy(
            TestPaths.Fixture("Stations/Station2_Report_Shift_3_12-May-2026.csv"),
            Path.Combine(env.S2Folder, "Station2_Report_Shift_1_01-Jan-2026.csv"));

        var service = CreateService(env, createDummy: true);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedExcelReportAsync(slot);

        Assert.NotNull(result);
        Assert.Contains("S2", result!.DummyStationsUsed);
        Assert.True(Directory.GetFiles(env.S2Folder, "*_DUMMY.csv").Length > 0);
    }

    [Fact]
    public async Task GenerateCombinedExcelReportAsync_normalized_csv_uses_exact_gocator_headers()
    {
        using var env = await TestDataSeeder.SeedFullEnvironmentAsync();
        var service = CreateService(env);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedExcelReportAsync(slot);
        Assert.NotNull(result);

        var text = await File.ReadAllTextAsync(result!.NormalizedCsvPath!);
        Assert.Contains("BLB2_B1PushBack", text, StringComparison.Ordinal);
        Assert.DoesNotContain(",Shift,Gocator,", text, StringComparison.Ordinal);
        Assert.DoesNotContain(",Top:Date,Gocator,", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task GenerateCombinedExcelReportAsync_sends_row_count_mismatch_alert_when_spread_exceeds_threshold()
    {
        using var env = await TestDataSeeder.SeedFullEnvironmentAsync();
        // Fixture sources (raw Gocator Top/Bottom + all 4 stations) match at 2 rows each — inflate Station 1 well
        // past the default threshold (20) so the min/max spread genuinely exceeds it rather than testing an edge case.
        InflateStationRowCount(Path.Combine(env.S1Folder, "Station1_Report_Shift_3_12-May-2026.csv"), extraRows: 25);

        var rowCountMismatchAlerts = new Mock<IRowCountMismatchAlertSender>();
        rowCountMismatchAlerts
            .Setup(s => s.SendRowCountMismatchAlertAsync(
                It.IsAny<ReportSlotContext>(),
                It.IsAny<IReadOnlyList<(string, int)>>(),
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        var service = CreateService(env, rowCountMismatchAlerts: rowCountMismatchAlerts.Object);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedExcelReportAsync(slot);

        Assert.NotNull(result);
        Assert.NotEmpty(result!.ExcelFilePath);
        // Every count here comes straight from a raw file — no merged Gocator number in the mix.
        rowCountMismatchAlerts.Verify(
            s => s.SendRowCountMismatchAlertAsync(
                It.Is<ReportSlotContext>(c => c.Shift == "3"),
                It.Is<IReadOnlyList<(string, int)>>(counts =>
                    counts.Count == 6 &&
                    counts.Single(c => c.Item1 == "Gocator Top").Item2 == 2 &&
                    counts.Single(c => c.Item1 == "Gocator Bottom").Item2 == 2 &&
                    counts.Single(c => c.Item1 == "Station 1 (S1)").Item2 == 27 &&
                    counts.Single(c => c.Item1 == "Station 2 (S2)").Item2 == 2 &&
                    counts.Single(c => c.Item1 == "Station 4 (S4)").Item2 == 2 &&
                    counts.Single(c => c.Item1 == "Station 5 (S5)").Item2 == 2),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task GenerateCombinedExcelReportAsync_does_not_send_row_count_mismatch_alert_when_disabled()
    {
        using var env = await TestDataSeeder.SeedFullEnvironmentAsync();
        InflateStationRowCount(Path.Combine(env.S1Folder, "Station1_Report_Shift_3_12-May-2026.csv"), extraRows: 25);

        var rowCountMismatchAlerts = new Mock<IRowCountMismatchAlertSender>();

        // Threshold <= 0 disables the check entirely, regardless of how large the actual row-count gap is.
        var service = CreateService(env, emailOptions: new EmailOptions { RowCountMismatchThreshold = 0 },
            rowCountMismatchAlerts: rowCountMismatchAlerts.Object);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedExcelReportAsync(slot);

        Assert.NotNull(result);
        rowCountMismatchAlerts.Verify(
            s => s.SendRowCountMismatchAlertAsync(
                It.IsAny<ReportSlotContext>(),
                It.IsAny<IReadOnlyList<(string, int)>>(),
                It.IsAny<CancellationToken>()),
            Times.Never);
    }

    /// <summary>Appends duplicate data rows to a station CSV fixture so its row count diverges from Gocator's.</summary>
    private static void InflateStationRowCount(string filePath, int extraRows)
    {
        var lines = File.ReadAllLines(filePath);
        var lastDataLine = lines[^1];
        var appended = Enumerable.Repeat(lastDataLine, extraRows);
        File.AppendAllLines(filePath, appended);
    }

    private static CombinedExcelReportService CreateService(
        ExportTestEnvironment env,
        bool? createDummy = null,
        EmailOptions? emailOptions = null,
        IRowCountMismatchAlertSender? rowCountMismatchAlerts = null) =>
        new(
            env.ExportPathsOptions(createDummy ?? false),
            Options.Create(emailOptions ?? new EmailOptions()),
            env.PathResolver,
            new GocatorCsvMergeService(
                env.ExportPathsOptions(createDummy ?? false),
                Options.Create(new EmailOptions()),
                env.PathResolver,
                NoOpSlottedAlertCoordinator.Instance,
                NullLogger<GocatorCsvMergeService>.Instance,
                env.Configuration),
            new StationDummyShiftCsvService(),
            NoOpSlottedAlertCoordinator.Instance,
            rowCountMismatchAlerts ?? NoOpRowCountMismatchAlertSender.Instance,
            NullLogger<CombinedExcelReportService>.Instance);
}
