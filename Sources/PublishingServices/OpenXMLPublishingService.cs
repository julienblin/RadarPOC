﻿using System;
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
                using (WordprocessingDocument package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
                {
                    var mainDoc = package.AddMainDocumentPart();

                    var altChunk = new AltChunk();
                    altChunk.Id = "AltChunkId1";

                    var altChunkPart = package.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml, altChunk.Id);
                    using (var altChunkStream = altChunkPart.GetStream())
                    using (var stringStream = new StreamWriter(altChunkStream))
                    {
                        stringStream.Write("<html><head/><body>" + opDoc.Discussion + "</body></html>");
                    }

                    package.MainDocumentPart.Document =
                        new Document(
                            new Body(
                                new Paragraph(
                                    new Run(
                                        new Text("")
                                    ),
                                    altChunk
                                )
                            )
                        );

                    SetPackageProperties(package);
                }

                result = stream.ToArray();
            }

            return result;
        }

        private void SetPackageProperties(OpenXmlPackage package)
        {
            package.PackageProperties.Creator = "ppelletier";
            package.PackageProperties.Title = "Product to review";
            package.PackageProperties.Subject = "";
            package.PackageProperties.Keywords = "";
            package.PackageProperties.Description = "";
            package.PackageProperties.Revision = "2";
            package.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2010-01-14T15:51:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2010-01-14T15:51:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            package.PackageProperties.LastModifiedBy = "Julien Blin";
            package.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2006-09-26T13:33:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }
    }
}