using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace Russell.RADAR.POC.PublishingServices
{
    public class WPMLPublishingService : IPublishingService
    {
        private string templateDirectory;

        public WPMLPublishingService(string templateDirectory)
        {
            this.templateDirectory = templateDirectory;
        }

        public byte[] Publish(Document document)
        {
            var opinionDoc = (OpinionDocument)document;

            var templateFile = Path.Combine(templateDirectory, @"OpinionDocument.xml");
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(templateFile);

            var nsmgr = new XmlNamespaceManager(new NameTable());
            nsmgr.AddNamespace("w", "http://schemas.microsoft.com/office/word/2003/wordml");
            nsmgr.AddNamespace("wx", "http://schemas.microsoft.com/office/word/2003/auxHint");
            nsmgr.AddNamespace("wsp", "http://schemas.microsoft.com/office/word/2003/wordml/sp2");

            var wxSubsection = xmlDoc.SelectSingleNode("/w:wordDocument/w:body/wx:sect/wx:sub-section", nsmgr);

            var topicHeaderP = wxSubsection.SelectSingleNode("w:p[@wsp:rsidR='00F8047A']", nsmgr);

            var topicHeaderText = topicHeaderP.SelectSingleNode("w:r[@wsp:rsidRPr='00233025']/w:t", nsmgr);
            topicHeaderText.InnerText = @"DISCUSSION";

            var topicContentP = wxSubsection.SelectSingleNode("w:p[@wsp:rsidR='00385C36']", nsmgr);
            var topicContentText = topicContentP.SelectSingleNode("w:r/w:t", nsmgr);

            var newTopicContent = xmlDoc.CreateElement("w:p", "http://schemas.microsoft.com/office/word/2003/wordml");
                newTopicContent.InnerXml = TransformXHTMLContent(opinionDoc.Discussion); ;
            topicContentText.ParentNode.ParentNode.ReplaceChild(newTopicContent, topicContentText.ParentNode);

            var utf8Encoding = new UTF8Encoding();

            return utf8Encoding.GetBytes(xmlDoc.OuterXml);
        }

        private string TransformXHTMLContent(string xhtmlContent)
        {
            var reader = new StringReader("<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\"><p>" + xhtmlContent + "</p>");
            var xmlReader = new XmlTextReader(reader);
            xmlReader.XmlResolver = new XmlUrlResolver();
            xmlReader.EntityHandling = EntityHandling.ExpandEntities;

            XslCompiledTransform myXslTrans = new XslCompiledTransform();
            myXslTrans.Load(Path.Combine(templateDirectory, @"XHTML2WordML.xslt"));
            
            var writer = new StringWriter();
            myXslTrans.Transform(xmlReader, null, writer);

            return writer.ToString();
        }
    }
}
