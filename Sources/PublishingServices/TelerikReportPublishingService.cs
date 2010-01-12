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

            var reportProcessor = new ReportProcessor();
            var renderResult = reportProcessor.RenderReport("PDF", template, null);
            return renderResult.DocumentBytes;
        }
    }
}
