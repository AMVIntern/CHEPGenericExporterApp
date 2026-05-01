namespace CHEPGenericExporterApp.Models;

/// <summary>Outcome of Gocator merge: output path when successful, and whether a slotted missing-file alert was already sent.</summary>
public readonly record struct GocatorMergeAttemptResult(string? CombinedCsvPath, bool SentSlottedMissingFileAlert);
