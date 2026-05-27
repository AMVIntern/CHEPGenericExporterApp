using System.Globalization;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace CHEPGenericExporterApp.Services;

public sealed class CsvAuditLogger
{
    private const string Header =
        "Shift,Date,GocatorReportSent,CombinedReportSent,LastAttemptUtc,MissingAlertCount,MissingAlertFinalized";

    private static readonly string[] ExpectedHeaderLegacy =
        ["Shift", "Date", "GocatorReportSent", "CombinedReportSent", "LastAttemptUtc"];

    private static readonly string[] ExpectedHeaderFull =
        ["Shift", "Date", "GocatorReportSent", "CombinedReportSent", "LastAttemptUtc", "MissingAlertCount", "MissingAlertFinalized"];

    private readonly string _logFilePath;
    private readonly object _sync = new();

    public CsvAuditLogger(IConfiguration configuration)
    {
        var configuredPath = configuration["LogFilePath"]?.Trim();
        if (string.IsNullOrWhiteSpace(configuredPath))
            throw new InvalidOperationException("Configuration key 'LogFilePath' is required and cannot be empty.");

        var fullPath = Path.GetFullPath(configuredPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            throw new InvalidOperationException(
                $"Configuration key 'LogFilePath' points to an invalid directory: '{configuredPath}'.");

        _logFilePath = fullPath;
        EnsureFileExists();
    }

    public string LogFilePath => _logFilePath;

    public int GetMissedSendCheckIntervalMinutes(IConfiguration configuration)
    {
        var raw = configuration["MissedSendCheckIntervalMinutes"];
        if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var minutes) || minutes <= 0)
            throw new InvalidOperationException(
                "Configuration key 'MissedSendCheckIntervalMinutes' is required and must be a positive integer.");

        return minutes;
    }

    public void EnsureRow(string shift, DateOnly date)
    {
        var normalizedShift = NormalizeShift(shift);
        lock (_sync)
        {
            var rows = ReadAllRowsNoLock();
            if (rows.Any(r => r.Shift == normalizedShift && r.Date == date))
                return;

            rows.Add(new CsvAuditRow(normalizedShift, date, false, false, null, 0, false));
            WriteAllRowsNoLock(rows);
        }
    }

    public void MarkGocatorSent(string shift, DateOnly date)
    {
        UpdateRow(shift, date, (row, now) => row with { GocatorReportSent = true, LastAttemptUtc = now });
    }

    public void MarkCombinedSent(string shift, DateOnly date)
    {
        UpdateRow(shift, date, (row, now) => row with { CombinedReportSent = true, LastAttemptUtc = now });
    }

    /// <summary>Returns true when the audit already records a successful Gocator email for this slot.</summary>
    public bool IsGocatorSent(string shift, DateOnly date)
    {
        var normalizedShift = NormalizeShift(shift);
        lock (_sync)
        {
            var rows = ReadAllRowsNoLock();
            var row = rows.FirstOrDefault(r => r.Shift == normalizedShift && r.Date == date);
            return row.Shift is not null && row.GocatorReportSent;
        }
    }

    /// <summary>Returns true when the audit already records a successful Combined email for this slot.</summary>
    public bool IsCombinedSent(string shift, DateOnly date)
    {
        var normalizedShift = NormalizeShift(shift);
        lock (_sync)
        {
            var rows = ReadAllRowsNoLock();
            var row = rows.FirstOrDefault(r => r.Shift == normalizedShift && r.Date == date);
            return row.Shift is not null && row.CombinedReportSent;
        }
    }

    public void MarkAttempt(string shift, DateOnly date)
    {
        UpdateRow(shift, date, (row, now) => row with { LastAttemptUtc = now });
    }

    /// <summary>Whether a missing-input-file alert may still be sent: true while <c>MissingAlertCount</c> is below <paramref name="maxAlerts"/> and the row is not finalized.</summary>
    public bool CanSendMissingFileAlert(string shift, DateOnly date, int maxAlerts)
    {
        if (maxAlerts <= 0)
            return true;

        var normalizedShift = NormalizeShift(shift);
        lock (_sync)
        {
            var rows = ReadAllRowsNoLock();
            var row = rows.FirstOrDefault(r => r.Shift == normalizedShift && r.Date == date);
            if (row.Shift is null)
                return true;

            return !row.MissingAlertFinalized && row.MissingAlertCount < maxAlerts;
        }
    }

    /// <summary>Call after a missing-file alert was delivered successfully: increments <c>MissingAlertCount</c>; sets <c>MissingAlertFinalized</c> when count reaches <paramref name="maxAlerts"/> (same value as Email:MaxMissingFileAlertsPerShiftDate).</summary>
    public void RecordMissingFileAlertSent(string shift, DateOnly date, int maxAlerts)
    {
        if (maxAlerts <= 0)
            return;

        var normalizedShift = NormalizeShift(shift);
        lock (_sync)
        {
            var rows = ReadAllRowsNoLock();
            var index = rows.FindIndex(r => r.Shift == normalizedShift && r.Date == date);
            if (index < 0)
            {
                var newCount = 1;
                rows.Add(new CsvAuditRow(normalizedShift, date, false, false, null, newCount, newCount >= maxAlerts));
            }
            else
            {
                var row = rows[index];
                var newCount = row.MissingAlertCount + 1;
                rows[index] = row with
                {
                    MissingAlertCount = newCount,
                    MissingAlertFinalized = newCount >= maxAlerts
                };
            }

            WriteAllRowsNoLock(rows);
        }
    }

    public IReadOnlyList<CsvAuditRow> GetPendingRows()
    {
        lock (_sync)
        {
            return ReadAllRowsNoLock()
                .Where(r => !r.GocatorReportSent || !r.CombinedReportSent)
                .ToList();
        }
    }

    private void UpdateRow(string shift, DateOnly date, Func<CsvAuditRow, DateTimeOffset, CsvAuditRow> updater)
    {
        var normalizedShift = NormalizeShift(shift);
        lock (_sync)
        {
            var rows = ReadAllRowsNoLock();
            var index = rows.FindIndex(r => r.Shift == normalizedShift && r.Date == date);
            if (index < 0)
                return;

            rows[index] = updater(rows[index], DateTimeOffset.UtcNow);
            WriteAllRowsNoLock(rows);
        }
    }

    private void EnsureFileExists()
    {
        lock (_sync)
        {
            if (!File.Exists(_logFilePath))
                File.WriteAllText(_logFilePath, Header + Environment.NewLine, Encoding.UTF8);
        }
    }

    private List<CsvAuditRow> ReadAllRowsNoLock()
    {
        EnsureFileExists();
        var lines = File.ReadAllLines(_logFilePath);
        if (lines.Length == 0)
            return [];

        var headerCols = lines[0].Split(',', StringSplitOptions.None);
        var extended = HeaderMatches(headerCols, ExpectedHeaderFull);
        var legacy = HeaderMatches(headerCols, ExpectedHeaderLegacy);

        if (!extended && !legacy)
        {
            throw new InvalidOperationException(
                $"CSV header in '{_logFilePath}' is invalid. Expected legacy ({string.Join(",", ExpectedHeaderLegacy)}) or full ({string.Join(",", ExpectedHeaderFull)}).");
        }

        var rows = new List<CsvAuditRow>();
        foreach (var rawLine in lines.Skip(1))
        {
            var line = rawLine.Trim();
            if (line.Length == 0)
                continue;

            var columns = line.Split(',', StringSplitOptions.None);
            if (columns.Length < ExpectedHeaderLegacy.Length)
                continue;

            var shift = NormalizeShift(columns[0]);
            if (!DateOnly.TryParseExact(columns[1].Trim(), "dd-MMM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                continue;

            var gocatorSent = bool.TryParse(columns[2].Trim(), out var g) && g;
            var combinedSent = bool.TryParse(columns[3].Trim(), out var c) && c;

            DateTimeOffset? lastAttemptUtc = null;
            var attemptRaw = columns[4].Trim();
            if (attemptRaw.Length > 0 &&
                DateTimeOffset.TryParse(attemptRaw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed))
            {
                lastAttemptUtc = parsed.ToUniversalTime();
            }

            var missingCount = 0;
            var missingFinalized = false;
            if (extended && columns.Length >= ExpectedHeaderFull.Length)
            {
                _ = int.TryParse(columns[5].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out missingCount);
                missingFinalized = bool.TryParse(columns[6].Trim(), out var mf) && mf;
            }

            rows.Add(new CsvAuditRow(shift, date, gocatorSent, combinedSent, lastAttemptUtc, missingCount, missingFinalized));
        }

        return rows;
    }

    private void WriteAllRowsNoLock(IReadOnlyList<CsvAuditRow> rows)
    {
        var lines = new List<string>(rows.Count + 1) { Header };
        foreach (var row in rows.OrderBy(r => r.Date).ThenBy(r => r.Shift, StringComparer.Ordinal))
        {
            lines.Add(string.Join(",",
                row.Shift,
                row.Date.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture),
                row.GocatorReportSent ? "true" : "false",
                row.CombinedReportSent ? "true" : "false",
                row.LastAttemptUtc?.ToString("O", CultureInfo.InvariantCulture) ?? string.Empty,
                row.MissingAlertCount.ToString(CultureInfo.InvariantCulture),
                row.MissingAlertFinalized ? "true" : "false"));
        }

        File.WriteAllLines(_logFilePath, lines, Encoding.UTF8);
    }

    private static string NormalizeShift(string shift)
    {
        var normalized = shift.Trim();
        if (!int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out var shiftNumber) ||
            shiftNumber is < 1 or > 3)
        {
            throw new InvalidOperationException($"Invalid shift '{shift}'. Expected 1, 2, or 3.");
        }

        return shiftNumber.ToString(CultureInfo.InvariantCulture);
    }

    private static bool HeaderMatches(string[] columns, string[] expected)
    {
        if (columns.Length != expected.Length)
            return false;
        for (var i = 0; i < expected.Length; i++)
        {
            if (!string.Equals(columns[i].Trim(), expected[i], StringComparison.Ordinal))
                return false;
        }

        return true;
    }
}

public readonly record struct CsvAuditRow(
    string Shift,
    DateOnly Date,
    bool GocatorReportSent,
    bool CombinedReportSent,
    DateTimeOffset? LastAttemptUtc,
    int MissingAlertCount,
    bool MissingAlertFinalized);
