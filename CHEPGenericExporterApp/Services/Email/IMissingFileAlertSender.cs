using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Services.Email;

/// <summary>Sends internal operational emails when expected CSV inputs are missing.</summary>
public interface IMissingFileAlertSender
{
    /// <summary>Sends one alert email listing all <paramref name="missingDescriptions"/> (non-empty lines).</summary>
    /// <param name="scheduledSlot">When set, subject is tied to shift/date (no unrelated-day fallbacks).</param>
    Task SendMissingFilesAlertAsync(
        IReadOnlyList<string> missingDescriptions,
        CancellationToken cancellationToken = default,
        ReportSlotContext? scheduledSlot = null);
}
