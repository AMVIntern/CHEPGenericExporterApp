namespace CHEPGenericExporterApp.Configuration;

/// <summary>Mailbox, recipients, and templates. <see cref="Password"/> is used as the SMTP password when <c>smtp:Password</c> is not set (same account as <see cref="FromAddress"/>).</summary>
public sealed class EmailOptions
{
    public const string SectionName = "Email";

    public string FromAddress { get; set; } = "";

    /// <summary>SMTP / app password for the mailbox in <see cref="FromAddress"/> (e.g. Gmail app password).</summary>
    public string Password { get; set; } = "";

    public List<string> ToAddresses { get; set; } = new();
    public List<string> CcAddresses { get; set; } = new();

    /// <summary>Internal recipients for missing raw CSV / input file alerts (not customer report mail).</summary>
    public List<string> InternalAmvTeam { get; set; } = new();

    /// <summary>
    /// When true, Gocator and combined customer emails are sent whenever the report file exists,
    /// without requiring the attachment name to match the scheduled shift/date. When false, mismatches are blocked and internal alerts are sent.
    /// </summary>
    public bool BypassReportAttachmentSlotCheck { get; set; }

    /// <summary>
    /// Per shift + report date: each successful internal missing-file alert increments <c>MissingAlertCount</c> in the audit CSV;
    /// <c>MissingAlertFinalized</c> is set to true when that count reaches this value (e.g. 3 allows three delivered alerts, then no more for that slot).
    /// This is separate from SMTP <c>smtp:MaxRetryAttempts</c> (retries inside one send). Set to 0 or less to disable the cap.
    /// </summary>
    public int MaxMissingFileAlertsPerShiftDate { get; set; } = 3;

    /// <summary>
    /// When false, the Gocator CSV merge step does not send internal missing-file emails (Top/Bottom/merge failures).
    /// The combined Excel step still sends one alert listing all issues, including Top/Bottom. Use when the scheduler runs Gocator then Combined so you do not get two emails per slot.
    /// When true (default), merge sends as today; set false only if a combined report job always follows for the same slot.
    /// </summary>
    public bool SendInternalMissingFileAlertFromGocatorMerge { get; set; } = true;

    /// <summary>Internal missing-file alert subject when <c>scheduledSlot</c> is set. Placeholders: <c>{shift}</c>, <c>{date}</c> (report date as dd-MMM-yyyy).</summary>
    public string MissingFileAlertSubjectWithSlotTemplate { get; set; } =
        "Missing some files – Shift {shift}, {date}";

    /// <summary>Internal missing-file subject when there is no slot and exactly one issue line. Placeholder: <c>{detail}</c> (truncated).</summary>
    public string MissingFileAlertSubjectSingleIssueTemplate { get; set; } = "File is missing – {detail}";

    /// <summary>Internal missing-file subject when there is no slot and multiple issues. Placeholder: <c>{count}</c>.</summary>
    public string MissingFileAlertSubjectMultiIssueTemplate { get; set; } = "File is missing – {count} issues";

    public string CombinedReportSubjectTemplate { get; set; } =
        "AMV Combined Report - Shift {shift} - {date}";

    public string CombinedReportBodyWithZip { get; set; } = "";
    public string CombinedReportBodyWithoutZip { get; set; } = "";

    /// <summary>
    /// Appended to the standard internal missing-file alert (same subject as other missing-file mail) when dummy station CSVs were used.
    /// Placeholders: <c>{site}</c>, <c>{shift}</c>, <c>{date}</c>, <c>{stations}</c>, <c>{details}</c>.
    /// </summary>
    public string DummyStationGeneratedInternalAlertBodyTemplate { get; set; } =
        "Dummy station placeholder CSV(s) were auto-generated (zero values, *_DUMMY.csv) and the customer combined report was still sent.\r\nSite: {site}\r\nShift: {shift}\r\nReport date: {date}\r\nStations: {stations}\r\n{details}";

    /// <summary>Gocator-only subject (if you add a Gocator mail step).</summary>
    public string GocatorReportSubjectTemplate { get; set; } = "AMV Gocator Report";

    /// <summary>Gocator-only body; <c>{0}</c> = date, <c>{1}</c> = shift.</summary>
    public string GocatorReportBodyTemplate { get; set; } =
        "Please find attached the Gocator Report for {0} corresponding to Shift {1}.";

    /// <summary>
    /// When the spread between the lowest and highest row count among Gocator and stations 1/2/4/5 reaches at
    /// least this many rows, an internal row-count-mismatch alert is sent to <see cref="InternalAmvTeam"/>.
    /// Set to 0 or less to disable this check.
    /// </summary>
    public int RowCountMismatchThreshold { get; set; } = 20;

    /// <summary>Row-count-mismatch alert subject. Placeholders: <c>{shift}</c>, <c>{date}</c>.</summary>
    public string RowCountMismatchAlertSubjectTemplate { get; set; } =
        "Row count mismatch – Shift {shift}, {date}";

    /// <summary>
    /// Row-count-mismatch alert summary line (min vs max source). Placeholders: <c>{diff}</c>, <c>{shift}</c>,
    /// <c>{date}</c>, <c>{minSource}</c>, <c>{minCount}</c>, <c>{maxSource}</c>, <c>{maxCount}</c>.
    /// </summary>
    public string RowCountMismatchAlertSummaryTemplate { get; set; } =
        "Row count difference of {diff} for Shift {shift}, Date {date}: lowest is {minSource} ({minCount} rows), highest is {maxSource} ({maxCount} rows).";

    /// <summary>
    /// Row-count-mismatch alert body, one line per data source (Gocator, S1, S2, S4, S5). Placeholders:
    /// <c>{source}</c>, <c>{count}</c>, <c>{shift}</c>, <c>{date}</c>.
    /// </summary>
    public string RowCountMismatchAlertLineTemplate { get; set; } = "{source}: {count} rows";
}
