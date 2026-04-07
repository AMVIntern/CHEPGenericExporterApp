using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Configuration;

/// <summary>Fills missing <see cref="SmtpOptions.UserName"/> / <see cref="SmtpOptions.Password"/> from <see cref="EmailOptions"/> when not set under <c>Smtp</c>.</summary>
public sealed class SmtpFromEmailConfigurator : IConfigureOptions<SmtpOptions>
{
    private readonly IOptions<EmailOptions> _email;

    public SmtpFromEmailConfigurator(IOptions<EmailOptions> email)
    {
        _email = email;
    }

    public void Configure(SmtpOptions options)
    {
        var e = _email.Value;

        if (string.IsNullOrWhiteSpace(options.UserName) && !string.IsNullOrWhiteSpace(e.FromAddress))
            options.UserName = e.FromAddress.Trim();

        if (string.IsNullOrWhiteSpace(options.Password) && !string.IsNullOrWhiteSpace(e.Password))
            options.Password = e.Password.Trim();
    }
}
