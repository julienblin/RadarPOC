using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ent = Russell.RADAR.POC.Entities;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Russell.RADAR.POC.PublishingServices
{
    public class OpenXMLPublishingService : IPublishingService
    {
        public byte[] Publish(Ent.Document document)
        {
            var opDoc = (Ent.OpinionDocument)document;
            byte[] result = null;
            using (var stream = new MemoryStream())
            {
                OpinionDocumentGenerator docGenerator = new OpinionDocumentGenerator();
                docGenerator.CreatePackage(stream, opDoc);

                result = stream.ToArray();
            }

            return result;
        }
    }
}
