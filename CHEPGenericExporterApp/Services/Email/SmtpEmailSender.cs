using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Extensions.Options;
using MimeKit;

namespace CHEPGenericExporterApp.Services.Email;

public sealed class SmtpEmailSender : IEmailSender
{
    private readonly SmtpOptions _smtp;
    private readonly ILogger<SmtpEmailSender> _logger;

    public SmtpEmailSender(IOptions<SmtpOptions> smtpOptions, ILogger<SmtpEmailSender> logger)
    {
        _smtp = smtpOptions.Value;
        _logger = logger;
    }

    public async Task<bool> SendAsync(OutgoingEmail message, CancellationToken cancellationToken = default)
    {
        if (message.To.Count == 0)
        {
            _logger.LogWarning("Email send skipped: no recipients.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(_smtp.Password))
        {
            _logger.LogWarning(
                "Email send skipped: set smtp:Password (or smtp:smtp_password) or Email:Password in appsettings, user secrets, or environment (SMTP requires authentication).");
            return false;
        }

        if (!string.IsNullOrEmpty(message.PrimaryAttachmentPath) && !File.Exists(message.PrimaryAttachmentPath))
        {
            _logger.LogWarning("Email send skipped: primary attachment not found at {Path}.", message.PrimaryAttachmentPath);
            return false;
        }

        if (message.AdditionalAttachmentPaths != null)
        {
            foreach (var path in message.AdditionalAttachmentPaths)
            {
                if (!string.IsNullOrEmpty(path) && !File.Exists(path))
                {
                    _logger.LogWarning("Email send skipped: additional attachment not found at {Path}.", path);
                    return false;
                }
            }
        }

        var secureSocket = ResolveSecureSocketOptions();

        for (var attempt = 1; attempt <= _smtp.MaxRetryAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                using var mime = BuildMimeMessage(message);

                using var smtpClient = new SmtpClient();
                await smtpClient
                    .ConnectAsync(_smtp.Host, _smtp.Port, secureSocket, cancellationToken)
                    .ConfigureAwait(false);

                await smtpClient
                    .AuthenticateAsync(_smtp.UserName, _smtp.Password, cancellationToken)
                    .ConfigureAwait(false);

                await smtpClient.SendAsync(mime, cancellationToken).ConfigureAwait(false);
                await smtpClient.DisconnectAsync(true, cancellationToken).ConfigureAwait(false);

                _logger.LogInformation("Email sent successfully to {Recipients}.", string.Join(", ", message.To));
                return true;
            }
            catch (MailKit.ServiceNotAuthenticatedException ex)
            {
                _logger.LogWarning(ex, "SMTP attempt {Attempt}/{Max} failed: authentication rejected ({Message}).",
                    attempt, _smtp.MaxRetryAttempts, ex.Message);
                if (attempt < _smtp.MaxRetryAttempts)
                    await Task.Delay(_smtp.DelayBetweenRetriesMilliseconds, cancellationToken).ConfigureAwait(false);
            }
            catch (MailKit.ProtocolException ex)
            {
                _logger.LogWarning(ex, "SMTP attempt {Attempt}/{Max} failed: {Message}.",
                    attempt, _smtp.MaxRetryAttempts, ex.Message);
                if (attempt < _smtp.MaxRetryAttempts)
                    await Task.Delay(_smtp.DelayBetweenRetriesMilliseconds, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                _logger.LogWarning(ex, "Email attempt {Attempt}/{Max} failed: {Message}.",
                    attempt, _smtp.MaxRetryAttempts, ex.Message);
                if (attempt < _smtp.MaxRetryAttempts)
                    await Task.Delay(_smtp.DelayBetweenRetriesMilliseconds, cancellationToken).ConfigureAwait(false);
            }
        }

        _logger.LogError("Email send failed after {Max} attempts.", _smtp.MaxRetryAttempts);
        return false;
    }

    private SecureSocketOptions ResolveSecureSocketOptions()
    {
        if (!_smtp.EnableSsl)
            return SecureSocketOptions.None;

        // Port 465 uses implicit TLS; 587 and typical submission ports use STARTTLS (fixes MustIssueStartTlsFirst with O365/Gmail).
        if (_smtp.Port == 465)
            return SecureSocketOptions.SslOnConnect;

        return SecureSocketOptions.StartTls;
    }

    private static MimeMessage BuildMimeMessage(OutgoingEmail message)
    {
        var mime = new MimeMessage();
        mime.From.Add(MailboxAddress.Parse(message.From));
        foreach (var to in message.To)
            mime.To.Add(MailboxAddress.Parse(to));
        if (message.Cc != null)
        {
            foreach (var cc in message.Cc)
                mime.Cc.Add(MailboxAddress.Parse(cc));
        }

        mime.Subject = message.Subject;

        var builder = new BodyBuilder { TextBody = message.Body };

        if (!string.IsNullOrEmpty(message.PrimaryAttachmentPath) && File.Exists(message.PrimaryAttachmentPath))
            builder.Attachments.Add(message.PrimaryAttachmentPath);

        if (message.AdditionalAttachmentPaths != null)
        {
            foreach (var path in message.AdditionalAttachmentPaths)
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    builder.Attachments.Add(path);
            }
        }

        mime.Body = builder.ToMessageBody();
        return mime;
    }
}
