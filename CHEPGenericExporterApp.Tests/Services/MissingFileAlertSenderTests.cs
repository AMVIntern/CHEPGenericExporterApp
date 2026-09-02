using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.Models;
using CHEPGenericExporterApp.Services;
using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Scheduling;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Moq;

namespace CHEPGenericExporterApp.Tests.Services;

/// <summary>
/// Covers the weekend suppression rule added to <see cref="MissingFileAlertSender"/>: alerts must be
/// suppressed for Saturday (any shift) and for Sunday's Shift 1/2, but NOT for Sunday's Shift 3 — which
/// is produced by Monday's slot-0 job and is required to keep its pre-existing, unsuppressed behaviour.
/// Every <see cref="ReportSlotContext"/> in the codebase carries the true business Shift/Date (scheduler,
/// recovery worker, and filename parsing alike), so the rule is exercised directly on (Shift, ReportDate)
/// pairs rather than any notion of "which day the job fired on".
/// </summary>
public sealed class MissingFileAlertSenderTests
{
    private static MissingFileAlertSender CreateSender(Mock<IEmailSender> emailSender)
    {
        var options = Options.Create(new EmailOptions
        {
            FromAddress = "amv@example.com",
            InternalAmvTeam = new List<string> { "team@example.com" },
            MaxMissingFileAlertsPerShiftDate = 0 // disable the per-slot cap; not what's under test here
        });

        var auditPath = Path.Combine(Path.GetTempPath(), $"audit_{Guid.NewGuid():N}.csv");
        var config = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?> { ["LogFilePath"] = auditPath })
            .Build();
        var auditLogger = new CsvAuditLogger(config);

        return new MissingFileAlertSender(emailSender.Object, options, auditLogger, NullLogger<MissingFileAlertSender>.Instance);
    }

    private static async Task<bool> WasEmailSentAsync(string shift, DateOnly reportDate)
    {
        var emailSender = new Mock<IEmailSender>();
        emailSender
            .Setup(s => s.SendAsync(It.IsAny<OutgoingEmail>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var sender = CreateSender(emailSender);
        var ctx = new ReportSlotContext(shift, reportDate.ToString("dd-MMM-yyyy"), reportDate);

        await sender.SendMissingFilesAlertAsync(
            new[] { "TOP file missing." },
            CancellationToken.None,
            scheduledSlot: ctx,
            applyPerSlotMissingAlertLimit: false);

        return emailSender.Invocations.Count > 0;
    }

    // Saturday 2026-05-16, Sunday 2026-05-17, Monday 2026-05-18 (matches the week used in
    // ScheduleCalculatorTests.GetNextScheduledJob_produces_the_expected_shift_set_for_a_full_week).
    private static readonly DateOnly Saturday = new(2026, 5, 16);
    private static readonly DateOnly Sunday = new(2026, 5, 17);
    private static readonly DateOnly Monday = new(2026, 5, 18);

    [Theory]
    [InlineData("1")]
    [InlineData("2")]
    [InlineData("3")]
    public async Task Saturday_alerts_are_suppressed_for_every_shift(string shift)
    {
        Assert.False(await WasEmailSentAsync(shift, Saturday));
    }

    [Theory]
    [InlineData("1")]
    [InlineData("2")]
    public async Task Sunday_shift_1_and_2_alerts_are_suppressed(string shift)
    {
        Assert.False(await WasEmailSentAsync(shift, Sunday));
    }

    [Fact]
    public async Task Sunday_shift_3_alert_is_NOT_suppressed_because_it_is_already_handled_by_the_existing_system()
    {
        // ReportDate.DayOfWeek == Sunday, same as Shift 1/2 above — the shift number is what
        // must save this one from suppression. This is the case the original patch got wrong:
        // a suppression rule keyed only on ReportDate.DayOfWeek silently broke this alert.
        Assert.True(await WasEmailSentAsync("3", Sunday));
    }

    [Theory]
    [InlineData("1")]
    [InlineData("2")]
    [InlineData("3")]
    public async Task Weekday_alerts_are_never_suppressed(string shift)
    {
        Assert.True(await WasEmailSentAsync(shift, Monday));
    }

    [Fact]
    public async Task No_scheduled_slot_never_suppresses_ad_hoc_alerts()
    {
        var emailSender = new Mock<IEmailSender>();
        emailSender
            .Setup(s => s.SendAsync(It.IsAny<OutgoingEmail>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(true);

        var sender = CreateSender(emailSender);

        await sender.SendMissingFilesAlertAsync(
            new[] { "Some ad-hoc issue with no scheduled slot." },
            CancellationToken.None,
            scheduledSlot: null,
            applyPerSlotMissingAlertLimit: false);

        emailSender.Verify(s => s.SendAsync(It.IsAny<OutgoingEmail>(), It.IsAny<CancellationToken>()), Times.Once);
    }
}
