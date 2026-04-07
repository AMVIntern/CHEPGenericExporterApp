# CHEP Generic Exporter

Console application that periodically merges Gocator Top/Bottom CSV files, builds a combined Excel report (and optional normalized CSV/zip), and emails the result using SMTP settings from configuration.

## Running

1. Publish or build the project (`dotnet build` / `dotnet publish`).
2. Copy `appsettings.json` next to the executable (it is copied automatically on build).
3. Edit **`appsettings.json` only** on the target machine (all paths, schedule, mailbox, recipients, templates). Set `Email:Password` to your Gmail app password (or equivalent) on that machine; do not commit secrets to source control.
4. Run `CHEPGenericExporterApp.exe` (or `dotnet run` from the project folder).

## Configuration sections

| Section | Purpose |
|--------|---------|
| **Logging** | Standard .NET logging levels. |
| **Scheduler** | When to run the export (see below). |
| **Smtp** | SMTP server only: host, port, SSL, timeouts, retries. **User name and password are not set here**—they are taken from **Email** (`FromAddress` and `Password`). |
| **Email** | **Single place for the mailbox and recipients:** `FromAddress`, `Password` (SMTP/app password), `ToAddresses`, `CcAddresses`, Gocator templates (`GocatorReportSubjectTemplate`, `GocatorReportBodyTemplate` with `{0}`/`{1}`), and combined report templates (`CombinedReportSubjectTemplate`, `CombinedReportBodyWithZip`, `CombinedReportBodyWithoutZip` with `{shift}` / `{date}`). |
| **Imap** | Optional future IMAP ingestion (host, credentials); not used by the current pipeline. |
| **ExportPaths** | Folders for Top/Bottom CSVs, combined output, station logs, and normalized report site code. |

Paths under **ExportPaths** may be **absolute** or **relative to the application base directory** (the folder containing the executable).

### Scheduler (AMV shift timing)

Wall-clock times use **`Scheduler:TimeZoneId`** (e.g. `Australia/Sydney` or `Local`).

- **`GocatorTimes`** — When to run the Gocator Top/Bottom CSV merge (default **06:00**, **14:00**, **22:00**).
- **`CombinedTimes`** — When to run the combined Excel report + email (default **06:02**, **14:02**, **22:02**). Must have the **same number** of entries as `GocatorTimes` (each pair is one shift).

**Calendar rules:**

- **Saturday:** no runs.
- **Sunday:** only the **last** pair in the lists (default **22:00** Gocator, **22:02** combined).
- **Monday–Friday:** all pairs in order (Gocator first, then combined two minutes later by default).

**Other flags:**

- **`RunOnStart`** — Run the full pipeline once immediately, then follow the schedule.
- **`StatusLogIntervalSeconds`** — While waiting for the next run, log remaining wait at this interval.

## Secrets and portability

- Put the real **`Email:Password`** only in the deployed copy of `appsettings.json` (or use environment variables to override keys if you add them later). Do not commit production passwords to git.
- On a new PC, copy the app and edit **`appsettings.json`** only; no code changes are required.
