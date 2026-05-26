using CHEPGenericExporterApp.Services;

namespace CHEPGenericExporterApp.Tests.Services;

public sealed class ExportPathResolverTests
{
    [Fact]
    public void Resolve_returns_absolute_path_unchanged()
    {
        var resolver = new ExportPathResolver();
        var input = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "abs-test-folder"));
        Assert.Equal(input, resolver.Resolve(input));
    }

    [Fact]
    public void Resolve_roots_relative_path_under_base_directory()
    {
        var resolver = new ExportPathResolver();
        var resolved = resolver.Resolve("relative/sub");
        Assert.True(Path.IsPathFullyQualified(resolved));
        Assert.StartsWith(AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar), resolved, StringComparison.OrdinalIgnoreCase);
    }
}
