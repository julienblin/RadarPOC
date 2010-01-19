using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using HtmlAgilityPack;
using System.Web;

namespace Russell.RADAR.POC.Entities.Content
{
    public class FormattedContent : BaseFormattedElement
    {
        public void FromXHTML(string input)
        {
            Children.Clear();

            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(input);

            ParseRecursive(this, doc.DocumentNode);
        }

        private void ParseRecursive(IFormattedElement baseElement, HtmlNode baseNode)
        {
            foreach (var childNode in baseNode.ChildNodes)
            {
                IFormattedElement createdNode = null;
                switch (childNode.Name)
                {
                    case "p":
                        createdNode = new ParagraphFormattedElement();
                        break;
                    case "b":
                    case "strong":
                        createdNode = new BoldFormattedElement();
                        break;
                    case "i":
                    case "em":
                        createdNode = new ItalicFormattedContent();
                        break;
                    default:
                        createdNode = new TextFormattedElement(HttpUtility.HtmlDecode(childNode.InnerText));
                        break;
                }
                baseElement.Children.Add(createdNode);
                ParseRecursive(createdNode, childNode);
            }
        }

        public string ToXHTML()
        {
            var builder = new StringBuilder();
            ToXHTML(builder);
            return builder.ToString();
        }

        public override void ToXHTML(StringBuilder builder)
        {
            ForEachChild(x => x.ToXHTML(builder));
        }

        public override OpenXmlElement ToOpenXmlElement()
        {
            throw new NotImplementedException();
        }

        public IList<DocumentFormat.OpenXml.Wordprocessing.Paragraph> GetParagraphs()
        {
            var result = new List<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
            ForEachChild(x =>
            {
                if (x is ParagraphFormattedElement)
                    result.Add((DocumentFormat.OpenXml.Wordprocessing.Paragraph)x.ToOpenXmlElement());

            });
            return result;
        }

        public override object Clone()
        {
            var clone = new FormattedContent();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
