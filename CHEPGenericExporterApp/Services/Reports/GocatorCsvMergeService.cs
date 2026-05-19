using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;

namespace CHEPGenericExporterApp.Services.Reports;

/// <summary>Merges Top/Bottom Gocator CSV files into a combined report in the configured combined folder.</summary>
public class GocatorCsvMergeService
{
    private readonly ExportPathResolver _pathResolver;
    private readonly IMissingFileSlottedAlertCoordinator _missingFileAlerts;
    private readonly EmailOptions _email;
    private readonly ILogger<GocatorCsvMergeService> _logger;
    private readonly IConfiguration _configuration;

    public GocatorCsvMergeService(
        IOptions<ExportPathsOptions> pathsOptions,
        IOptions<EmailOptions> emailOptions,
        ExportPathResolver pathResolver,
        IMissingFileSlottedAlertCoordinator missingFileAlerts,
        ILogger<GocatorCsvMergeService> logger,
        IConfiguration configuration)
    {
        Paths = pathsOptions.Value;
        _email = emailOptions.Value;
        _pathResolver = pathResolver;
        _missingFileAlerts = missingFileAlerts;
        _logger = logger;
        _configuration = configuration;
    }

    private ExportPathsOptions Paths { get; }

    private bool SendSlottedMissingFromMerge => _email.SendInternalMissingFileAlertFromGocatorMerge;

    private async Task<bool> TrySendMergeMissingSlottedAsync(
        ReportSlotContext slot,
        IReadOnlyList<string> lines,
        CancellationToken cancellationToken)
    {
        if (!SendSlottedMissingFromMerge)
            return false;

        await _missingFileAlerts.SendOrEnqueueSlottedAsync(
                slot,
                applyPerSlotMissingAlertLimit: true,
                lines,
                cancellationToken)
            .ConfigureAwait(false);
        return true;
    }

    /// <summary>
    /// Appends the same Top/Bottom raw-input checks used by merge (folder / values CSV) into <paramref name="missing"/> for a single combined missing-file email from the Excel step.
    /// </summary>
    public void AppendTopBottomInputMissingLines(ReportSlotContext slot, IList<string> missing, DateTime? targetDate = null)
    {
        var effectiveDate = targetDate ?? slot.ReportDate.ToDateTime(TimeOnly.MinValue);
        if (missing is List<string> list)
            AddTopBottomInputIssues(slot, list, effectiveDate);
        else
        {
            var temp = new List<string>();
            AddTopBottomInputIssues(slot, temp, effectiveDate);
            foreach (var line in temp)
                missing.Add(line);
        }
    }

    /// <summary>
    /// Merges Top/Bottom *values* CSVs for the scheduled shift/date (filename tokens when present), matches rows within 1.5s, writes
    /// <c>Gocator_Report_Shift_{shift}_{date}.csv</c>. Returns no path when inputs or merged shift/date do not match the slot.
    /// </summary>
    public virtual async Task<GocatorMergeAttemptResult> GenerateCombinedCsvAsync(
        ReportSlotContext slot,
        CancellationToken cancellationToken = default,
        DateTime? targetDate = null)
    {
        try
        {
            var topFolder = _pathResolver.Resolve(Paths.TopCsvFolder);
            var bottomFolder = _pathResolver.Resolve(Paths.BottomCsvFolder);
            var combinedFolder = _pathResolver.Resolve(Paths.GocatorCombinedFolder);

             Directory.CreateDirectory(combinedFolder);

            var effectiveDate = targetDate ?? slot.ReportDate.ToDateTime(TimeOnly.MinValue);
            var missing = new List<string>();
            AddTopBottomInputIssues(slot, missing, effectiveDate);

            if (missing.Count > 0)
            {
                var sent = await TrySendMergeMissingSlottedAsync(slot, missing, cancellationToken).ConfigureAwait(false);
                return new GocatorMergeAttemptResult(null, SentSlottedMissingFileAlert: sent);
            }

            var topFile = FindLatestValuesCsv(topFolder, effectiveDate, slot.Shift)!;
            var bottomFile = FindLatestValuesCsv(bottomFolder, effectiveDate, slot.Shift)!;

            var topData = ReadCsvFile(topFile, "Top");
            if (topData == null || topData.Rows.Count == 0)
            {
                _logger.LogWarning("Top CSV file has insufficient data rows.");
                var sentTop = await TrySendMergeMissingSlottedAsync(
                    slot,
                    new[] { $"Top Gocator raw CSV is unreadable or has no data rows: {topFile}" },
                    cancellationToken).ConfigureAwait(false);
                return new GocatorMergeAttemptResult(null, SentSlottedMissingFileAlert: sentTop);
            }

            var bottomData = ReadCsvFile(bottomFile, "Bottom");
            if (bottomData == null || bottomData.Rows.Count == 0)
            {
                _logger.LogWarning("Bottom CSV file has insufficient data rows.");
                var sentBot = await TrySendMergeMissingSlottedAsync(
                    slot,
                    new[] { $"Bottom Gocator raw CSV is unreadable or has no data rows: {bottomFile}" },
                    cancellationToken).ConfigureAwait(false);
                return new GocatorMergeAttemptResult(null, SentSlottedMissingFileAlert: sentBot);
            }

            string topDateCol = FindColumnByName(topData.Headers, new[] { "top:date" });
            string topTimestampCol = FindColumnByName(topData.Headers, new[] { "top:timestamp" });
            string bottomDateCol = FindColumnByName(bottomData.Headers, new[] { "bot:date" });
            string bottomTimestampCol = FindColumnByName(bottomData.Headers, new[] { "bot:timestamp" });

            if (string.IsNullOrEmpty(topDateCol) || string.IsNullOrEmpty(topTimestampCol) ||
                string.IsNullOrEmpty(bottomDateCol) || string.IsNullOrEmpty(bottomTimestampCol))
            {
                _logger.LogWarning("Could not find required date/timestamp columns in CSV files.");
                var sentCols = await TrySendMergeMissingSlottedAsync(
                    slot,
                    new[]
                    {
                        "Could not find required date/timestamp columns in Top/Bottom CSV (need top:date, top:timestamp, bot:date, bot:timestamp)."
                    },
                    cancellationToken).ConfigureAwait(false);
                return new GocatorMergeAttemptResult(null, SentSlottedMissingFileAlert: sentCols);
            }

            CalculateFullTimestamps(topData.Rows, topDateCol, topTimestampCol);
            CalculateFullTimestamps(bottomData.Rows, bottomDateCol, bottomTimestampCol);

            topData.Rows = topData.Rows.OrderBy(r => r.FullTimestamp).ToList();
            bottomData.Rows = bottomData.Rows.OrderBy(r => r.FullTimestamp).ToList();

            string topOverallCol = FindColumnByName(topData.Headers, new[] { "top:overall pass" }, true);
            string bottomOverallCol = FindColumnByName(bottomData.Headers, new[] { "bot:overall_result" }, true);

            var combinedHeaders = new List<string>(topData.Headers);
            string[] excludedColumns = { "Bot:Date", "Bot:Timestamp" };
            foreach (var header in bottomData.Headers)
            {
                if (!combinedHeaders.Contains(header, StringComparer.OrdinalIgnoreCase) &&
                    !excludedColumns.Contains(header, StringComparer.OrdinalIgnoreCase))
                {
                    combinedHeaders.Add(header);
                }
            }

            combinedHeaders.Add("Assured_Result");

            List<Dictionary<string, string>> combinedRows = new();

            int i = 0, j = 0;
            while (i < topData.Rows.Count && j < bottomData.Rows.Count)
            {
                if (!topData.Rows[i].FullTimestamp.HasValue || !bottomData.Rows[j].FullTimestamp.HasValue)
                {
                    i++;
                    j++;
                    continue;
                }

                double diff = Math.Abs((topData.Rows[i].FullTimestamp!.Value - bottomData.Rows[j].FullTimestamp!.Value).TotalSeconds);
                if (diff < 1.5)
                {
                    var combinedRow = new Dictionary<string, string>(topData.Rows[i].Data, StringComparer.OrdinalIgnoreCase);

                    foreach (var kvp in bottomData.Rows[j].Data)
                    {
                        if (!combinedRow.ContainsKey(kvp.Key))
                            combinedRow[kvp.Key] = kvp.Value;
                    }

                    if (!string.IsNullOrEmpty(topOverallCol) && !string.IsNullOrEmpty(bottomOverallCol) &&
                        topData.Rows[i].Data.ContainsKey(topOverallCol) && bottomData.Rows[j].Data.ContainsKey(bottomOverallCol))
                    {
                        if (double.TryParse(topData.Rows[i].Data[topOverallCol], NumberStyles.Any, CultureInfo.InvariantCulture, out double topVal) &&
                            double.TryParse(bottomData.Rows[j].Data[bottomOverallCol], NumberStyles.Any, CultureInfo.InvariantCulture, out double bottomVal))
                        {
                            combinedRow["Assured_Result"] = (topVal * bottomVal).ToString(CultureInfo.InvariantCulture);
                        }
                    }
                    else
                    {
                        combinedRow["Assured_Result"] = "";
                    }

                    combinedRows.Add(combinedRow);
                    i++;
                    j++;
                }
                else if (topData.Rows[i].FullTimestamp < bottomData.Rows[j].FullTimestamp)
                {
                    i++;
                }
                else
                {
                    j++;
                }
            }

            string shiftCol = FindColumnByName(combinedHeaders.ToArray(), new[] { "shift" });
            string dateCol = FindColumnByName(combinedHeaders.ToArray(), new[] { "top:date" });

            if (combinedRows.Count == 0)
            {
                _logger.LogWarning("No matching rows found between Top and Bottom CSV files.");
                var sentRows = await TrySendMergeMissingSlottedAsync(
                    slot,
                    new[]
                    {
                        "No matching rows found between Top and Bottom Gocator CSV files (within 1.5s timestamp pairing)."
                    },
                    cancellationToken).ConfigureAwait(false);
                return new GocatorMergeAttemptResult(null, SentSlottedMissingFileAlert: sentRows);
            }

            string shiftValue = shiftCol != null && combinedRows[0].ContainsKey(shiftCol)
                ? combinedRows[0][shiftCol]
                : "Unknown";
            string dateValue = dateCol != null && combinedRows[0].ContainsKey(dateCol)
                ? combinedRows[0][dateCol]
                : DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);

            var _siteCode = _configuration.GetSection("ExportPaths")["NormalizedReportSiteCode"];
            string combinedFile = Path.Combine(combinedFolder, $"{_siteCode}_Gocator_Report_Shift_{shiftValue}_{dateValue}.csv");

            using (var writer = new StreamWriter(combinedFile))
            {
                writer.WriteLine(string.Join(",", combinedHeaders));
                foreach (var row in combinedRows)
                {
                    var rowValues = combinedHeaders.Select(h => row.ContainsKey(h) ? row[h] : "").ToArray();
                    writer.WriteLine(string.Join(",", rowValues));
                }
            }

            if (!ReportCsvDate.ReportFileMatchesSlot(combinedFile, slot))
            {
                _logger.LogWarning(
                    "Merged Gocator CSV does not match scheduled Shift {Shift}, Date {Date} (file: {File}).",
                    slot.Shift,
                    slot.ReportDateDdMmmYyyy,
                    Path.GetFileName(combinedFile));
                var sentMismatch = await TrySendMergeMissingSlottedAsync(
                    slot,
                    new[]
                    {
                        $"Gocator merge data does not match scheduled Shift {slot.Shift}, Date {slot.ReportDateDdMmmYyyy} (produced {Path.GetFileName(combinedFile)})."
                    },
                    cancellationToken).ConfigureAwait(false);
                return new GocatorMergeAttemptResult(null, SentSlottedMissingFileAlert: sentMismatch);
            }

            _logger.LogInformation("Combined Gocator CSV saved to {Path}.", combinedFile);
            return new GocatorMergeAttemptResult(combinedFile, SentSlottedMissingFileAlert: false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating Gocator combined CSV.");
            return new GocatorMergeAttemptResult(null, SentSlottedMissingFileAlert: false);
        }
    }

    private void AddTopBottomInputIssues(ReportSlotContext slot, List<string> missing, DateTime effectiveDate)
    {
        var topFolder = _pathResolver.Resolve(Paths.TopCsvFolder);
        var bottomFolder = _pathResolver.Resolve(Paths.BottomCsvFolder);

        if (!Directory.Exists(topFolder))
            missing.Add($"Top Gocator folder missing: {topFolder}");
        else if (FindLatestValuesCsv(topFolder, effectiveDate, slot.Shift) == null)
        {
            _logger.LogWarning("No CSV file containing 'values' found in Top folder {Folder}.", topFolder);
            missing.Add($"No CSV file containing 'values' found in Top folder: {topFolder}");
        }

        if (!Directory.Exists(bottomFolder))
            missing.Add($"Bottom Gocator folder missing: {bottomFolder}");
        else if (FindLatestValuesCsv(bottomFolder, effectiveDate, slot.Shift) == null)
        {
            _logger.LogWarning("No CSV file containing 'values' found in Bottom folder {Folder}.", bottomFolder);
            missing.Add($"No CSV file containing 'values' found in Bottom folder: {bottomFolder}");
        }
    }

    private static string? FindLatestValuesCsv(string folder, DateTime targetDate, string? targetShift = null)
    {
        if (!Directory.Exists(folder))
            return null;

        var valuesFiles = Directory.GetFiles(folder, "*.csv")
            .Where(f => Path.GetFileName(f).ToLowerInvariant().Contains("values"));

        var tokenCandidates = new[]
        {
            targetDate.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture),
            targetDate.ToString("d-MMM-yyyy", CultureInfo.InvariantCulture),
            targetDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            targetDate.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture),
            targetDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture),
            targetDate.ToString("ddMMyyyy", CultureInfo.InvariantCulture)
        };

        var dateMatches = valuesFiles
            .Where(f =>
            {
                var name = Path.GetFileNameWithoutExtension(f);
                return tokenCandidates.Any(token => name.Contains(token, StringComparison.OrdinalIgnoreCase));
            });

        if (int.TryParse(targetShift, NumberStyles.Integer, CultureInfo.InvariantCulture, out var shiftNo))
        {
            var shiftTokens = new[]
            {
                $"shift_{shiftNo}",
                $"shift-{shiftNo}",
                $"shift{shiftNo}",
                $"s{shiftNo}_",
                $"_s{shiftNo}"
            };

            var dateAndShiftMatch = dateMatches
                .Where(f =>
                {
                    var name = Path.GetFileNameWithoutExtension(f);
                    return shiftTokens.Any(token => name.Contains(token, StringComparison.OrdinalIgnoreCase));
                })
                .OrderByDescending(f => File.GetLastWriteTime(f))
                .FirstOrDefault();

            return dateAndShiftMatch;
        }

        return null;
    }

    private CsvData? ReadCsvFile(string filePath, string sourceName)
    {
        try
        {
            string[] lines = File.ReadAllLines(filePath);
            if (lines.Length < 2)
            {
                _logger.LogWarning("{Source} CSV file has insufficient data rows.", sourceName);
                return null;
            }

            string[] headers = lines[0].Split(',').Select(h => h.Trim()).ToArray();
            var rows = new List<CsvRow>();

            for (int i = 1; i < lines.Length; i++)
            {
                string[] values = lines[i].Split(',');
                if (values.Length != headers.Length)
                {
                    _logger.LogDebug("Row {Row} in {Source} CSV column count mismatch. Skipping.", i, sourceName);
                    continue;
                }

                var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int j = 0; j < headers.Length; j++)
                    rowData[headers[j]] = values[j].Trim();

                rows.Add(new CsvRow { Data = rowData });
            }

            return new CsvData { Headers = headers, Rows = rows };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error reading {Source} CSV file.", sourceName);
            return null;
        }
    }

    private static string? FindColumnByName(string[] headers, string[] possibleNames, bool containsMatch = false)
    {
        foreach (var header in headers)
        {
            string headerLower = header.ToLowerInvariant();
            foreach (var name in possibleNames)
            {
                if (containsMatch)
                {
                    if (headerLower.Contains(name.ToLowerInvariant()))
                        return header;
                }
                else
                {
                    if (headerLower.Equals(name.ToLowerInvariant()) || headerLower.Replace(":", "").Equals(name.ToLowerInvariant()))
                        return header;
                }
            }
        }

        return null;
    }

    private static string? FindColumnByName(List<string> headers, string[] possibleNames, bool containsMatch = false) =>
        FindColumnByName(headers.ToArray(), possibleNames, containsMatch);

    private static void CalculateFullTimestamps(List<CsvRow> rows, string dateCol, string timestampCol)
    {
        foreach (var row in rows)
        {
            if (!row.Data.ContainsKey(dateCol) || !row.Data.ContainsKey(timestampCol))
                continue;

            string dateStr = row.Data[dateCol];
            string timeStr = row.Data[timestampCol].Trim();

            bool parsed = false;

            if (!DateTime.TryParseExact(dateStr, "dd-MMM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
            {
                if (!DateTime.TryParse(dateStr, out date))
                    continue;
            }

            if (timeStr.Contains("AM", StringComparison.OrdinalIgnoreCase) ||
                timeStr.Contains("PM", StringComparison.OrdinalIgnoreCase))
            {
                string[] formats12Hour =
                {
                    "h:mm:ss tt",
                    "hh:mm:ss tt",
                    "h:mm:ss.fff tt",
                    "hh:mm:ss.fff tt"
                };

                foreach (var format in formats12Hour)
                {
                    if (DateTime.TryParseExact(timeStr, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime timeOnly))
                    {
                        row.FullTimestamp = date.Date + timeOnly.TimeOfDay;
                        parsed = true;
                        break;
                    }
                }
            }
            else
            {
                if (TimeSpan.TryParseExact(timeStr, @"hh\:mm\:ss\.fff", CultureInfo.InvariantCulture, out TimeSpan time))
                {
                    row.FullTimestamp = date.Date + time;
                    parsed = true;
                }
                else if (TimeSpan.TryParseExact(timeStr, @"hh\:mm\:ss", CultureInfo.InvariantCulture, out time))
                {
                    row.FullTimestamp = date.Date + time;
                    parsed = true;
                }
            }

            if (!parsed && TimeSpan.TryParse(timeStr, out TimeSpan fallbackTime))
            {
                row.FullTimestamp = date.Date + fallbackTime;
                parsed = true;
            }

            if (!parsed)
            {
                string combinedDateTime = $"{dateStr} {timeStr}";
                if (DateTime.TryParse(combinedDateTime, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fullDateTime))
                    row.FullTimestamp = fullDateTime;
            }
        }
    }
}
