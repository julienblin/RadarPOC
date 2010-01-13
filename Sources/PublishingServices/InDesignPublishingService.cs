using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;
using System.IO;
using System.Reflection;

namespace Russell.RADAR.POC.PublishingServices
{
    public class InDesignPublishingService : IPublishingService
    {
        private string inDesignTemplateDirectory;

        public InDesignPublishingService(string inDesignTemplateDirectory)
        {
            this.inDesignTemplateDirectory = inDesignTemplateDirectory;
        }

        public ExportOption ExportOption
        {
            get { return ExportOption.AsFile; }
        }

        public byte[] PublishAsPDF(Document document)
        {
            throw new NotImplementedException();
        }

        public string PublishAsPDFFile(Document document)
        {
            // Lovely COM...
            var missing = Type.Missing;
            var app = (InDesign.Application)COMCreateObject("InDesign.Application");

            try
            {
                var templateFile = GetTemplateFilePath(@"OpinionDocument");
                var doc = (InDesign.Document)app.Open(templateFile, false);

                var firstPage = (InDesign.Page)doc.Pages[1];
                var discussionTextFrame = (InDesign.TextFrame)firstPage.TextFrames[2];
                discussionTextFrame.Contents = ((OpinionDocument)document).Discussion;

                doc.Export(InDesign.idExportFormat.idPDFType, @"C:\TestPDF.pdf", false, missing, missing, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }

            return @"C:\TestPDF.pdf";
        }

        private static object COMCreateObject(string sProgID)
        {
            // We get the type using just the ProgID
            Type oType = Type.GetTypeFromProgID(sProgID);
            if (oType != null)
            {
                return Activator.CreateInstance(oType);
            }

            return null;
        }

        private string GetTemplateFilePath(string fileNameWoExtension)
        {
            return Path.Combine(inDesignTemplateDirectory, string.Format("{0}.indd", fileNameWoExtension));
        }
    }
}
