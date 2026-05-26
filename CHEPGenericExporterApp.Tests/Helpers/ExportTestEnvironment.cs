using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Tests.Helpers;

internal sealed class ExportTestEnvironment : IDisposable
{
    public string Root { get; }
    public string TopFolder { get; }
    public string BottomFolder { get; }
    public string GocatorCombinedFolder { get; }
    public string S1Folder { get; }
    public string S2Folder { get; }
    public string S4Folder { get; }
    public string S5Folder { get; }
    public string CombinedOutputFolder { get; }
    public ExportPathResolver PathResolver { get; } = new();

    public ExportTestEnvironment()
    {
        Root = Path.Combine(Path.GetTempPath(), "CHEPExporterTests", Guid.NewGuid().ToString("N"));
        TopFolder = Path.Combine(Root, "Top");
        BottomFolder = Path.Combine(Root, "Bottom");
        GocatorCombinedFolder = Path.Combine(Root, "Combined");
        S1Folder = Path.Combine(Root, "S1");
        S2Folder = Path.Combine(Root, "S2");
        S4Folder = Path.Combine(Root, "S4");
        S5Folder = Path.Combine(Root, "S5");
        CombinedOutputFolder = Path.Combine(Root, "Reports");

        Directory.CreateDirectory(TopFolder);
        Directory.CreateDirectory(BottomFolder);
        Directory.CreateDirectory(GocatorCombinedFolder);
        Directory.CreateDirectory(S1Folder);
        Directory.CreateDirectory(S2Folder);
        Directory.CreateDirectory(S4Folder);
        Directory.CreateDirectory(S5Folder);
        Directory.CreateDirectory(CombinedOutputFolder);
    }

    public IOptions<ExportPathsOptions> ExportPathsOptions(bool createDummyWhenMissing = false) =>
        Options.Create(new ExportPathsOptions
        {
            TopCsvFolder = TopFolder,
            BottomCsvFolder = BottomFolder,
            GocatorCombinedFolder = GocatorCombinedFolder,
            S1Folder = S1Folder,
            S2Folder = S2Folder,
            S4Folder = S4Folder,
            S5Folder = S5Folder,
            CombinedReportOutputFolder = CombinedOutputFolder,
            NormalizedReportSiteCode = "TEST",
            CreateDummyStationShiftCsvWhenMissing = createDummyWhenMissing
        });

    public IConfiguration Configuration => new ConfigurationBuilder()
        .AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["ExportPaths:NormalizedReportSiteCode"] = "TEST"
        })
        .Build();

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(Root))
                Directory.Delete(Root, recursive: true);
        }
        catch
        {
            // Best-effort cleanup for temp test directories.
        }
    }
}
