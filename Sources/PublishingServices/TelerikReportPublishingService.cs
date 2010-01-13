using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;
using Russell.RADAR.POC.PublishingServices.TelerikTemplates;
using Telerik.Reporting.Processing;
using System.IO;

namespace Russell.RADAR.POC.PublishingServices
{
    public class TelerikReportPublishingService : IPublishingService
    {
        public byte[] PublishAsPDF(Document document)
        {
            var opDoc = (OpinionDocument)document;

            var template = new OpinionDocumentTemplate();
            template.Discussion = opDoc.Discussion;
            template.InvestementStaff = opDoc.InvestmentStaff;

            var reportProcessor = new ReportProcessor();
            var renderResult = reportProcessor.RenderReport("PDF", template, null);
            return renderResult.DocumentBytes;
        }

        public string PublishAsPDFFile(Document document)
        {
            throw new NotImplementedException();
        }

        public ExportOption ExportOption
        {
            get { return ExportOption.AsByte; }
        }
    }
}
