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

    public string CombinedReportSubjectTemplate { get; set; } =
        "AMV Combined Report - Shift {shift} - {date}";

    public string CombinedReportBodyWithZip { get; set; } = "";
    public string CombinedReportBodyWithoutZip { get; set; } = "";

    /// <summary>Gocator-only subject (if you add a Gocator mail step).</summary>
    public string GocatorReportSubjectTemplate { get; set; } = "AMV Gocator Report";

    /// <summary>Gocator-only body; <c>{0}</c> = date, <c>{1}</c> = shift.</summary>
    public string GocatorReportBodyTemplate { get; set; } =
        "Please find attached the Gocator Report for {0} corresponding to Shift {1}.";
}
