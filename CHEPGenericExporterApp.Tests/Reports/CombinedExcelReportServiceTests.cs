using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Services.Reports;
using CHEPGenericExporterApp.Services.Scheduling;
using CHEPGenericExporterApp.Tests.Helpers;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

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
    public async Task GenerateCombinedExcelReportAsync_returns_null_when_gocator_missing()
    {
        using var env = new ExportTestEnvironment();
        TestDataSeeder.CopyStationFiles(env);
        var service = CreateService(env);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedExcelReportAsync(slot);

        Assert.Null(result);
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

    private static CombinedExcelReportService CreateService(ExportTestEnvironment env, bool? createDummy = null) =>
        new(
            env.ExportPathsOptions(createDummy ?? false),
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
            NullLogger<CombinedExcelReportService>.Instance);
}
