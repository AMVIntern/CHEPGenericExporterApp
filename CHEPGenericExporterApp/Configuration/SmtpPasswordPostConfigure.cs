using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Configuration;

/// <summary>
/// Maps <c>smtp:smtp_password</c> to <see cref="SmtpOptions.Password"/> when <c>smtp:Password</c> did not bind.
/// </summary>
public sealed class SmtpPasswordPostConfigure : IPostConfigureOptions<SmtpOptions>
{
    private readonly IConfiguration _configuration;

    public SmtpPasswordPostConfigure(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    public void PostConfigure(string? name, SmtpOptions options)
    {
        if (!string.IsNullOrWhiteSpace(options.Password))
            return;

        var raw = _configuration.GetSection(SmtpOptions.SectionName)["smtp_password"]
            ?? _configuration.GetSection("Smtp")["smtp_password"];
        if (string.IsNullOrWhiteSpace(raw))
            return;

        options.Password = raw.Trim();
    }
}
