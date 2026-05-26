using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services.Scheduling;

namespace CHEPGenericExporterApp.Services.Reports;

/// <summary>Builds placeholder station shift CSVs from an existing file's header rows in the same folder.</summary>
public sealed class StationDummyShiftCsvService
{
    private static readonly Regex ShiftReportTailRegex = new(
        @"Shift_(\d+)_(.+)$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex DummyFileNameRegex = new(
        @"_DUMMY$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// Writes <c>{prefix}_Shift_{n}_{date}_DUMMY.csv</c> with the same two header lines as a template file in <paramref name="stationFolder"/>.
    /// </summary>
    public bool TryCreateDummyShiftFile(
        string stationFolder,
        ReportSlotContext slot,
        CsvData gocatorData,
        string stationCode,
        string stationDisplayName,
        out string? dummyFilePath,
        out string? failureReason)
    {
        dummyFilePath = null;
        failureReason = null;

        if (!Directory.Exists(stationFolder))
        {
            failureReason = $"Station folder does not exist: {stationFolder}";
            return false;
        }

        if (gocatorData.Rows.Count == 0)
        {
            failureReason = "Gocator data has no rows to align dummy station timestamps with.";
            return false;
        }

        if (!TryLoadHeaderTemplate(stationFolder, out var headerLine, out var subHeaderLine, out var fileNamePrefix, out var templateStationValue))
        {
            failureReason = $"No template station CSV with header rows found in {stationFolder}.";
            return false;
        }

        string[] headers = headerLine.Split(',').Select(h => h.Trim()).ToArray();
        if (headers.Length < 4)
        {
            failureReason = "Template station CSV header row is too short.";
            return false;
        }

        string dateCol = FindColumn(gocatorData.Headers, "top:date");
        string timestampCol = FindColumn(gocatorData.Headers, "top:timestamp");
        string shiftCol = FindColumn(gocatorData.Headers, "shift");

        Directory.CreateDirectory(stationFolder);
        dummyFilePath = Path.Combine(
            stationFolder,
            $"{fileNamePrefix}_Shift_{slot.Shift}_{slot.ReportDateDdMmmYyyy}_DUMMY.csv");

        using (var writer = new StreamWriter(dummyFilePath, false, Encoding.UTF8))
        {
            writer.WriteLine(headerLine);
            writer.WriteLine(subHeaderLine);

            foreach (var gocatorRow in gocatorData.Rows)
            {
                var values = new string[headers.Length];
                for (int j = 0; j < headers.Length; j++)
                    values[j] = "0";

                values[0] = GetGocatorCell(gocatorRow, dateCol);
                values[1] = GetGocatorCell(gocatorRow, timestampCol);
                values[2] = GetGocatorCell(gocatorRow, shiftCol, slot.Shift);
                values[3] = templateStationValue;

                if (string.IsNullOrEmpty(values[0]))
                    values[0] = FormatStationDate(slot.ReportDate);
                if (string.IsNullOrEmpty(values[2]))
                    values[2] = slot.Shift;

                writer.WriteLine(string.Join(",", values));
            }
        }

        return true;
    }

    private static bool TryLoadHeaderTemplate(
        string stationFolder,
        out string headerLine,
        out string subHeaderLine,
        out string fileNamePrefix,
        out string templateStationValue)
    {
        headerLine = "";
        subHeaderLine = "";
        fileNamePrefix = "";
        templateStationValue = "";

        string? bestPath = null;
        var bestWrite = DateTime.MinValue;
        foreach (var path in Directory.GetFiles(stationFolder, "*.csv"))
        {
            var name = Path.GetFileNameWithoutExtension(path);
            if (DummyFileNameRegex.IsMatch(name))
                continue;
            if (!ShiftReportTailRegex.IsMatch(name))
                continue;

            var wt = File.GetLastWriteTime(path);
            if (wt >= bestWrite)
            {
                bestWrite = wt;
                bestPath = path;
            }
        }

        if (bestPath == null)
            return false;

        string[] lines = File.ReadAllLines(bestPath);
        if (lines.Length < 2)
            return false;

        headerLine = lines[0];
        subHeaderLine = lines[1];
        fileNamePrefix = ShiftReportTailRegex.Replace(Path.GetFileNameWithoutExtension(bestPath), "").TrimEnd('_');
        if (string.IsNullOrEmpty(fileNamePrefix))
            return false;

        if (lines.Length >= 3)
        {
            string[] sample = lines[2].Split(',');
            if (sample.Length >= 4)
                templateStationValue = sample[3].Trim();
        }

        if (string.IsNullOrEmpty(templateStationValue))
        {
            string[] headerParts = headerLine.Split(',');
            if (headerParts.Length >= 4 && headerParts[3].Equals("Station", StringComparison.OrdinalIgnoreCase))
                templateStationValue = "Station";
        }

        return true;
    }

    private static string FormatStationDate(DateOnly date) =>
        date.ToString("dd-MMM-yy", CultureInfo.InvariantCulture);

    private static string? FindColumn(string[] headers, string name)
    {
        foreach (var h in headers)
        {
            if (h.Equals(name, StringComparison.OrdinalIgnoreCase))
                return h;
        }

        return null;
    }

    private static string GetGocatorCell(CsvRow row, string? column, string fallback = "")
    {
        if (string.IsNullOrEmpty(column) || !row.Data.TryGetValue(column, out var v))
            return fallback;
        return v?.Trim() ?? fallback;
    }
}
