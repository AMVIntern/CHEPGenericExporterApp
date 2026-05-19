using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services;
using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Reports;
using CHEPGenericExporterApp.Services.Scheduling;
using CHEPGenericExporterApp.Tests.Helpers;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Moq;

namespace CHEPGenericExporterApp.Tests.Services;

public sealed class ExportPipelineTests
{
    [Fact]
    public async Task TryRecoverCombinedEmailAsync_skips_customer_email_when_attachment_slot_mismatch()
    {
        var ctx = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));
        var wrongPath = Path.Combine(Path.GetTempPath(), "TEST_Combined_Report_Shift_2_12-MAY-2026.xlsx");
        await File.WriteAllBytesAsync(wrongPath, new byte[] { 1 });

        try
        {
            var emailSender = new Mock<IEmailSender>(MockBehavior.Strict);
            var missingAlerts = new Mock<IMissingFileAlertSender>();
            missingAlerts
                .Setup(s => s.SendMissingFilesAlertAsync(
                    It.IsAny<IReadOnlyList<string>>(),
                    It.IsAny<CancellationToken>(),
                    ctx,
                    It.IsAny<bool>()))
                .Returns(Task.CompletedTask);

            var excel = CreateExcelServiceMock();
            excel
                .Setup(s => s.GenerateCombinedExcelReportAsync(ctx, It.IsAny<CancellationToken>()))
                .ReturnsAsync(new CombinedReportResult
                {
                    ExcelFilePath = wrongPath,
                    SiteCode = "TEST"
                });

            var pipeline = CreatePipeline(emailSender.Object, missingAlerts.Object, excel.Object, bypassSlotCheck: false);
            var sent = await pipeline.TryRecoverCombinedEmailAsync(ctx);

            Assert.False(sent);
            emailSender.Verify(s => s.SendAsync(It.IsAny<OutgoingEmail>(), It.IsAny<CancellationToken>()), Times.Never);
            missingAlerts.Verify(
                s => s.SendMissingFilesAlertAsync(
                    It.Is<IReadOnlyList<string>>(l => l.Any(x => x.Contains("Combined report is not for scheduled", StringComparison.OrdinalIgnoreCase))),
                    It.IsAny<CancellationToken>(),
                    ctx,
                    false),
                Times.Once);
        }
        finally
        {
            if (File.Exists(wrongPath))
                File.Delete(wrongPath);
        }
    }

    [Fact]
    public async Task TryRecoverCombinedEmailAsync_sends_customer_email_without_dummy_note()
    {
        var ctx = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));
        var path = Path.Combine(Path.GetTempPath(), "TEST_Combined_Report_Shift_3_12-MAY-2026.xlsx");
        await File.WriteAllBytesAsync(path, new byte[] { 1 });

        OutgoingEmail? captured = null;
        try
        {
            var emailSender = new Mock<IEmailSender>();
            emailSender
                .Setup(s => s.SendAsync(It.IsAny<OutgoingEmail>(), It.IsAny<CancellationToken>()))
                .Callback<OutgoingEmail, CancellationToken>((m, _) => captured = m)
                .ReturnsAsync(true);

            var missingAlerts = new Mock<IMissingFileAlertSender>();
            missingAlerts
                .Setup(s => s.SendMissingFilesAlertAsync(
                    It.IsAny<IReadOnlyList<string>>(),
                    It.IsAny<CancellationToken>(),
                    It.IsAny<ReportSlotContext?>(),
                    It.IsAny<bool>()))
                .Returns(Task.CompletedTask);

            var excel = CreateExcelServiceMock();
            excel
                .Setup(s => s.GenerateCombinedExcelReportAsync(ctx, It.IsAny<CancellationToken>()))
                .ReturnsAsync(new CombinedReportResult
                {
                    ExcelFilePath = path,
                    SiteCode = "TEST",
                    DummyStationsUsed = new[] { "S2" }
                });

            var pipeline = CreatePipeline(emailSender.Object, missingAlerts.Object, excel.Object, bypassSlotCheck: false);
            var sent = await pipeline.TryRecoverCombinedEmailAsync(ctx);

            Assert.True(sent);
            Assert.NotNull(captured);
            Assert.DoesNotContain("DUMMY", captured!.Body, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("placeholder", captured.Body, StringComparison.OrdinalIgnoreCase);

            missingAlerts.Verify(
                s => s.SendMissingFilesAlertAsync(
                    It.Is<IReadOnlyList<string>>(l => l.Any(x => x.Contains("S2", StringComparison.Ordinal))),
                    It.IsAny<CancellationToken>(),
                    ctx,
                    true),
                Times.Once);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public async Task TryRecoverGocatorEmailAsync_blocks_wrong_shift_attachment()
    {
        var ctx = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));
        var csvPath = Path.Combine(Path.GetTempPath(), "TEST_Gocator_Report_Shift_2_12-MAY-2026.csv");
        await File.WriteAllTextAsync(csvPath, "h");

        try
        {
            var emailSender = new Mock<IEmailSender>(MockBehavior.Strict);
            var missingAlerts = new Mock<IMissingFileAlertSender>();
            missingAlerts
                .Setup(s => s.SendMissingFilesAlertAsync(
                    It.IsAny<IReadOnlyList<string>>(),
                    It.IsAny<CancellationToken>(),
                    ctx,
                    It.IsAny<bool>()))
                .Returns(Task.CompletedTask);

            var merge = CreateGocatorMergeServiceMock();
            merge
                .Setup(s => s.GenerateCombinedCsvAsync(
                    ctx,
                    It.IsAny<CancellationToken>(),
                    It.IsAny<DateTime?>()))
                .ReturnsAsync(new GocatorMergeAttemptResult(csvPath, false));

            var pipeline = CreatePipeline(emailSender.Object, missingAlerts.Object, merge: merge.Object, bypassSlotCheck: false);
            var sent = await pipeline.TryRecoverGocatorEmailAsync(ctx);

            Assert.False(sent);
        }
        finally
        {
            if (File.Exists(csvPath))
                File.Delete(csvPath);
        }
    }

    private static Mock<CombinedExcelReportService> CreateExcelServiceMock()
    {
        using var env = new ExportTestEnvironment();
        return new Mock<CombinedExcelReportService>(
            env.ExportPathsOptions(),
            env.PathResolver,
            CreateGocatorMergeService(env),
            new StationDummyShiftCsvService(),
            NoOpSlottedAlertCoordinator.Instance,
            NullLogger<CombinedExcelReportService>.Instance);
    }

    private static Mock<GocatorCsvMergeService> CreateGocatorMergeServiceMock()
    {
        using var env = new ExportTestEnvironment();
        return new Mock<GocatorCsvMergeService>(
            env.ExportPathsOptions(),
            Options.Create(new EmailOptions()),
            env.PathResolver,
            NoOpSlottedAlertCoordinator.Instance,
            NullLogger<GocatorCsvMergeService>.Instance,
            env.Configuration);
    }

    private static GocatorCsvMergeService CreateGocatorMergeService(ExportTestEnvironment env) =>
        new(
            env.ExportPathsOptions(),
            Options.Create(new EmailOptions()),
            env.PathResolver,
            NoOpSlottedAlertCoordinator.Instance,
            NullLogger<GocatorCsvMergeService>.Instance,
            env.Configuration);

    private static ExportPipeline CreatePipeline(
        IEmailSender emailSender,
        IMissingFileAlertSender missingAlerts,
        CombinedExcelReportService? excel = null,
        GocatorCsvMergeService? merge = null,
        bool bypassSlotCheck = false)
    {
        using var env = new ExportTestEnvironment();
        excel ??= CreateExcelServiceMock().Object;
        merge ??= CreateGocatorMergeService(env);

        return new ExportPipeline(
            merge,
            excel,
            emailSender,
            new Mock<IScheduleCalculator>().Object,
            missingAlerts,
            new Mock<IMissingFileSlottedAlertCoordinator>().Object,
            new CsvAuditLogger(CreateAuditConfiguration()),
            Options.Create(new EmailOptions
            {
                FromAddress = "test@example.com",
                ToAddresses = new List<string> { "customer@example.com" },
                InternalAmvTeam = new List<string> { "internal@example.com" },
                BypassReportAttachmentSlotCheck = bypassSlotCheck,
                GocatorReportSubjectTemplate = "Gocator {shift} {date}",
                GocatorReportBodyTemplate = "Body {0} {1}",
                CombinedReportSubjectTemplate = "Combined {shift} {date}",
                CombinedReportBodyWithoutZip = "Standard body {shift} {date}",
                DummyStationGeneratedInternalAlertBodyTemplate = "Dummy info {stations}"
            }),
            NullLogger<ExportPipeline>.Instance);
    }

    private static IConfiguration CreateAuditConfiguration()
    {
        var auditPath = Path.Combine(Path.GetTempPath(), "CHEPExporterTests", $"audit_{Guid.NewGuid():N}.csv");
        Directory.CreateDirectory(Path.GetDirectoryName(auditPath)!);
        File.WriteAllText(
            auditPath,
            "Shift,Date,GocatorReportSent,CombinedReportSent,LastAttemptUtc,MissingAlertCount,MissingAlertFinalized");
        return new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?> { ["LogFilePath"] = auditPath })
            .Build();
    }
}
