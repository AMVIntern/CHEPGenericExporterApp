using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Services.Reports;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Tests.Helpers;

internal static class TestDataSeeder
{
    public static async Task<ExportTestEnvironment> SeedFullEnvironmentAsync(bool includeS2 = true)
    {
        var env = new ExportTestEnvironment();
        File.Copy(
            TestPaths.Fixture("Gocator/Top/values_shift_3_12-May-2026.csv"),
            Path.Combine(env.TopFolder, "values_shift_3_12-May-2026.csv"));
        File.Copy(
            TestPaths.Fixture("Gocator/Bottom/values_shift_3_12-May-2026.csv"),
            Path.Combine(env.BottomFolder, "values_shift_3_12-May-2026.csv"));

        var merge = new GocatorCsvMergeService(
            env.ExportPathsOptions(),
            Options.Create(new EmailOptions { SendInternalMissingFileAlertFromGocatorMerge = false }),
            env.PathResolver,
            NoOpSlottedAlertCoordinator.Instance,
            NullLogger<GocatorCsvMergeService>.Instance,
            env.Configuration);

        var slot = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));
        await merge.GenerateCombinedCsvAsync(slot);

        CopyStationFiles(env, includeS2);
        return env;
    }

    public static void CopyStationFiles(ExportTestEnvironment env, bool includeS2 = true)
    {
        File.Copy(
            TestPaths.Fixture("Stations/Station1_Report_Shift_3_12-May-2026.csv"),
            Path.Combine(env.S1Folder, "Station1_Report_Shift_3_12-May-2026.csv"));
        if (includeS2)
        {
            File.Copy(
                TestPaths.Fixture("Stations/Station2_Report_Shift_3_12-May-2026.csv"),
                Path.Combine(env.S2Folder, "Station2_Report_Shift_3_12-May-2026.csv"));
        }

        File.Copy(
            TestPaths.Fixture("Stations/Station4_Report_Shift_3_12-May-2026.csv"),
            Path.Combine(env.S4Folder, "Station4_Report_Shift_3_12-May-2026.csv"));
        File.Copy(
            TestPaths.Fixture("Stations/Station5_Report_Shift_3_12-May-2026.csv"),
            Path.Combine(env.S5Folder, "Station5_Report_Shift_3_12-May-2026.csv"));
    }
}
