namespace CHEPGenericExporterApp.Configuration;

/// <summary>
/// Reserved for future IMAP-based ingestion. The current pipeline reads local CSV folders only.
/// </summary>
public sealed class ImapOptions
{
    public const string SectionName = "Imap";

    public bool Enabled { get; set; }
    public string Host { get; set; } = "";
    public int Port { get; set; } = 993;
    public bool UseSsl { get; set; } = true;
    public string UserName { get; set; } = "";
    public string Password { get; set; } = "";
}
