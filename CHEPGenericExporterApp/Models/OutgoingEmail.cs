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
}
