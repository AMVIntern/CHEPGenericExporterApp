using CHEPGenericExporterApp.Services.Reports;
using CHEPGenericExporterApp.Services.Scheduling;
using CHEPGenericExporterApp.Tests.Helpers;
using Microsoft.Extensions.Logging.Abstractions;

namespace CHEPGenericExporterApp.Tests.Reports;

public sealed class GocatorCsvMergeServiceTests
{
    [Fact]
    public async Task GenerateCombinedCsvAsync_merges_top_bottom_for_scheduled_slot()
    {
        using var env = new ExportTestEnvironment();
        File.Copy(
            TestPaths.Fixture("Gocator/Top/values_shift_3_12-May-2026.csv"),
            Path.Combine(env.TopFolder, "values_shift_3_12-May-2026.csv"));
        File.Copy(
            TestPaths.Fixture("Gocator/Bottom/values_shift_3_12-May-2026.csv"),
            Path.Combine(env.BottomFolder, "values_shift_3_12-May-2026.csv"));

        var service = CreateService(env);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedCsvAsync(slot);

        Assert.False(result.SentSlottedMissingFileAlert);
        Assert.NotNull(result.CombinedCsvPath);
        Assert.True(File.Exists(result.CombinedCsvPath));
        Assert.True(ReportCsvDate.ReportFileMatchesSlot(result.CombinedCsvPath!, slot));

        var lines = File.ReadAllLines(result.CombinedCsvPath!);
        Assert.True(lines.Length >= 3);
        Assert.Contains("Assured_Result", lines[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GenerateCombinedCsvAsync_returns_null_when_top_folder_missing()
    {
        using var env = new ExportTestEnvironment();
        Directory.Delete(env.TopFolder, recursive: true);

        var service = CreateService(env);
        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var result = await service.GenerateCombinedCsvAsync(slot);

        Assert.Null(result.CombinedCsvPath);
    }

    private static GocatorCsvMergeService CreateService(ExportTestEnvironment env) =>
        new(
            env.ExportPathsOptions(),
            Microsoft.Extensions.Options.Options.Create(new CHEPGenericExporterApp.Configuration.EmailOptions()),
            env.PathResolver,
            NoOpSlottedAlertCoordinator.Instance,
            NullLogger<GocatorCsvMergeService>.Instance,
            env.Configuration);
}
