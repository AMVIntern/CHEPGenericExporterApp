using System.Globalization;
using CHEPGenericExporterApp.Models;

namespace CHEPGenericExporterApp.Tests.Helpers;

internal static class GocatorCsvFactory
{
    public static CsvData CreateMinimal(int rowCount = 2)
    {
        var headers = new[] { "Shift", "Top:Date", "Top:Timestamp", "Top:Overall Pass", "TLB2B1_PB" };
        var rows = new List<CsvRow>();
        for (int i = 0; i < rowCount; i++)
        {
            var ts = TimeSpan.FromSeconds(10 * i);
            rows.Add(new CsvRow
            {
                Data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Shift"] = "3",
                    ["Top:Date"] = "12-MAY-2026",
                    ["Top:Timestamp"] = ts.ToString(@"hh\:mm\:ss\.fff", CultureInfo.InvariantCulture),
                    ["Top:Overall Pass"] = "1.0",
                    ["TLB2B1_PB"] = "1.0"
                }
            });
        }

        return new CsvData { Headers = headers, Rows = rows };
    }
}
