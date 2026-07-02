namespace CHEPGenericExporterApp.Models;

public enum ReportOutcomeReason
{
    Success,
    NoCsvFile,
    SlotMismatch,
    NotConfigured,
    EmailFailed,
    ReportNotProduced
}
