namespace CHEPGenericExporterApp.Models;

/// <summary>Outcome of Gocator merge: output path when successful, whether a slotted missing-file alert was already sent, and the specific missing-input lines (Top/Bottom folder, CSV, columns, etc.) when it failed.</summary>
public readonly record struct GocatorMergeAttemptResult(
    string? CombinedCsvPath,
    bool SentSlottedMissingFileAlert,
    IReadOnlyList<string>? MissingInputs = null);
