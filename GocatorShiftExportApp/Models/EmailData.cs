using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GocatorShiftExportApp.Models
{
    public class EmailData
    {
        public string FromEmail { get; set; }
        public string AppPassword { get; set; }
        public List<string> ToEmails { get; set; }
        public List<string> CcEmails { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string AttachmentPath { get; set; }
        public List<string> AdditionalAttachmentPaths { get; set; }
    }

    public class CombinedReportResult
    {
        public string ExcelFilePath { get; set; }
        public string NormalizedCsvPath { get; set; }
        public string NormalizedZipPath { get; set; }
    }

    public class EmailConfig
    {
        public EmailData Settings { get; set; } // Renamed from EmailConfig to Settings
    }
}
