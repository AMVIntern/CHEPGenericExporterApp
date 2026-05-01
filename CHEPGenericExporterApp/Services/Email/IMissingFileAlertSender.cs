using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Services.Email;

/// <summary>Sends internal operational emails when expected CSV inputs are missing.</summary>
public interface IMissingFileAlertSender
{
    /// <summary>Sends one alert email listing all <paramref name="missingDescriptions"/> (non-empty lines).</summary>
    /// <param name="scheduledSlot">When set, subject uses <c>Email:MissingFileAlertSubjectWithSlotTemplate</c> (<c>{shift}</c>, <c>{date}</c>).</param>
    /// <param name="applyPerSlotMissingAlertLimit">When true with <paramref name="scheduledSlot"/>, uses audit counters (see <c>Email:MaxMissingFileAlertsPerShiftDate</c>).</param>
    Task SendMissingFilesAlertAsync(
        IReadOnlyList<string> missingDescriptions,
        CancellationToken cancellationToken = default,
        ReportSlotContext? scheduledSlot = null,
        bool applyPerSlotMissingAlertLimit = false);
}
