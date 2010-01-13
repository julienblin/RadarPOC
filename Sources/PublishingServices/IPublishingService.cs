using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;

namespace Russell.RADAR.POC.PublishingServices
{
    public interface IPublishingService
    {
        byte[] PublishAsPDF(Document document);
        string PublishAsPDFFile(Document document);

        ExportOption ExportOption { get; }
    }

    public enum ExportOption
    {
        AsByte,
        AsFile
    }
}
