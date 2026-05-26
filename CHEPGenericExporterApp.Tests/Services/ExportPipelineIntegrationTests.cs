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

public sealed class ExportPipelineIntegrationTests
{
    [Fact]
    public async Task TryRecoverCombinedEmailAsync_sends_when_full_environment_is_seeded()
    {
        using var env = await TestDataSeeder.SeedFullEnvironmentAsync();
        var pipeline = CreateRealPipeline(env);
        var ctx = new ReportSlotContext("3", "12-MAY-2026", new DateOnly(2026, 5, 12));

        var sent = await pipeline.TryRecoverCombinedEmailAsync(ctx);

        Assert.True(sent);
    }

    private static ExportPipeline CreateRealPipeline(ExportTestEnvironment env)
    {
        var emailSender = new Mock<IEmailSender>();
        emailSender
            .Setup(s => s.SendAsync(It.IsAny<OutgoingEmail>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var missingAlerts = new Mock<IMissingFileAlertSender>();
        missingAlerts
            .Setup(s => s.SendMissingFilesAlertAsync(
                It.IsAny<IReadOnlyList<string>>(),
                It.IsAny<CancellationToken>(),
                It.IsAny<ReportSlotContext?>(),
                It.IsAny<bool>()))
            .Returns(Task.CompletedTask);

        return new ExportPipeline(
            new GocatorCsvMergeService(
                env.ExportPathsOptions(),
                Options.Create(new EmailOptions { SendInternalMissingFileAlertFromGocatorMerge = false }),
                env.PathResolver,
                NoOpSlottedAlertCoordinator.Instance,
                NullLogger<GocatorCsvMergeService>.Instance,
                env.Configuration),
            new CombinedExcelReportService(
                env.ExportPathsOptions(),
                env.PathResolver,
                new GocatorCsvMergeService(
                    env.ExportPathsOptions(),
                    Options.Create(new EmailOptions()),
                    env.PathResolver,
                    NoOpSlottedAlertCoordinator.Instance,
                    NullLogger<GocatorCsvMergeService>.Instance,
                    env.Configuration),
                new StationDummyShiftCsvService(),
                NoOpSlottedAlertCoordinator.Instance,
                NullLogger<CombinedExcelReportService>.Instance),
            emailSender.Object,
            new ScheduleCalculator(Options.Create(new SchedulerOptions
            {
                TimeZoneId = "UTC",
                GocatorTimes = new List<string> { "06:00", "14:00", "22:00" },
                CombinedTimes = new List<string> { "06:02", "14:02", "22:02" }
            })),
            missingAlerts.Object,
            NoOpSlottedAlertCoordinator.Instance,
            new CsvAuditLogger(CreateAuditConfiguration()),
            Options.Create(new EmailOptions
            {
                FromAddress = "test@example.com",
                ToAddresses = new List<string> { "customer@example.com" },
                InternalAmvTeam = new List<string> { "internal@example.com" },
                BypassReportAttachmentSlotCheck = false,
                CombinedReportSubjectTemplate = "Combined {shift} {date}",
                CombinedReportBodyWithoutZip = "Body {shift} {date}",
                DummyStationGeneratedInternalAlertBodyTemplate = "Dummy {stations}"
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
