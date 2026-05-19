namespace CHEPGenericExporterApp.Tests.Helpers;

internal static class TestPaths
{
    public static string FixtureRoot =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "TestFixtures"));

    public static string Fixture(string relativePath) =>
        Path.Combine(FixtureRoot, relativePath);
}
