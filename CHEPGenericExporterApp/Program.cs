using CHEPGenericExporterApp.Configuration;
using CHEPGenericExporterApp.HostedServices;
using CHEPGenericExporterApp.Services;
using CHEPGenericExporterApp.Services.Email;
using CHEPGenericExporterApp.Services.Reports;
using CHEPGenericExporterApp.Services.Scheduling;

var builder = Host.CreateApplicationBuilder(args);

builder.Services.Configure<SchedulerOptions>(builder.Configuration.GetSection(SchedulerOptions.SectionName));
builder.Services.Configure<SmtpOptions>(builder.Configuration.GetSection(SmtpOptions.SectionName));
builder.Services.Configure<EmailOptions>(builder.Configuration.GetSection(EmailOptions.SectionName));
builder.Services.Configure<ExportPathsOptions>(builder.Configuration.GetSection(ExportPathsOptions.SectionName));
builder.Services.Configure<ImapOptions>(builder.Configuration.GetSection(ImapOptions.SectionName));

// SMTP: appsettings "smtp" section; Email fills gaps; smtp_password alias in SmtpPasswordPostConfigure.
builder.Services.AddSingleton<IConfigureOptions<SmtpOptions>, SmtpFromEmailConfigurator>();
builder.Services.AddSingleton<IPostConfigureOptions<SmtpOptions>, SmtpPasswordPostConfigure>();

builder.Services.AddSingleton<ExportPathResolver>();
builder.Services.AddSingleton<IScheduleCalculator, ScheduleCalculator>();
builder.Services.AddSingleton<SmtpEmailSender>();
builder.Services.AddSingleton<IEmailRetryQueue, InMemoryEmailRetryQueue>();
builder.Services.AddSingleton<IEmailSender, ReliableEmailSender>();
builder.Services.AddSingleton<IMissingFileAlertSender, MissingFileAlertSender>();
builder.Services.AddSingleton<GocatorCsvMergeService>();
builder.Services.AddSingleton<CombinedExcelReportService>();
builder.Services.AddSingleton<ExportPipeline>();
builder.Services.AddHostedService<ScheduledExportWorker>();
builder.Services.AddHostedService<EmailRetryWorker>();

var host = builder.Build();
await host.RunAsync();
