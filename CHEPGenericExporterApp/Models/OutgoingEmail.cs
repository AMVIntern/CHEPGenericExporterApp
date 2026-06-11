namespace CHEPGenericExporterApp.Models;

public sealed class OutgoingEmail
{
    public required string From { get; init; }
    public required IReadOnlyList<string> To { get; init; }
    public IReadOnlyList<string>? Cc { get; init; }
    public required string Subject { get; init; }
    public required string Body { get; init; }
    public string? PrimaryAttachmentPath { get; init; }
    public IReadOnlyList<string>? AdditionalAttachmentPaths { get; init; }

    /// <summary>
    /// When set, the EmailRetryWorker will mark this slot as sent in the audit after a
    /// successful background retry — preventing MissedEmailRecoveryWorker from re-sending.
    /// </summary>
    public EmailAuditContext? AuditContext { get; init; }
}

/// <summary>Identifies which audit row to mark after a background retry succeeds.</summary>
public sealed record EmailAuditContext(string Shift, DateOnly ReportDate, EmailAuditKind Kind);

public enum EmailAuditKind { Gocator, Combined }
