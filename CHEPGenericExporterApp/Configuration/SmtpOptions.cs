namespace CHEPGenericExporterApp.Configuration;

/// <summary>SMTP settings bound from the <c>smtp</c> configuration section.</summary>
public sealed class SmtpOptions
{
    /// <summary>Configuration section key in appsettings (e.g. <c>"smtp"</c> object).</summary>
    public const string SectionName = "smtp";

    public string Host { get; set; } = "smtp.gmail.com";
    public int Port { get; set; } = 587;
    public bool EnableSsl { get; set; } = true;

    /// <summary>SMTP login; set under <c>smtp:UserName</c> or left empty to use <c>Email:FromAddress</c>.</summary>
    public string UserName { get; set; } = "";

    /// <summary>SMTP password; <c>smtp:Password</c> or <c>smtp:smtp_password</c> (or <c>Email:Password</c> if empty).</summary>
    public string Password { get; set; } = "";
    public int TimeoutMilliseconds { get; set; } = 300_000;
    public int MaxRetryAttempts { get; set; } = 3;
    public int DelayBetweenRetriesMilliseconds { get; set; } = 5_000;
}
