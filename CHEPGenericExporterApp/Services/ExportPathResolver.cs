namespace CHEPGenericExporterApp.Services;

/// <summary>Maps configured paths to absolute paths (relative paths are rooted at the app base directory).</summary>
public sealed class ExportPathResolver
{
    public string Resolve(string configuredPath)
    {
        if (string.IsNullOrWhiteSpace(configuredPath))
            return configuredPath;

        var trimmed = configuredPath.Trim();
        if (Path.IsPathFullyQualified(trimmed))
            return Path.GetFullPath(trimmed);

        return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, trimmed));
    }
}
