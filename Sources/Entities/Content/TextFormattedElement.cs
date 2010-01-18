using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using System.Web;

namespace Russell.RADAR.POC.Entities.Content
{
    public class TextFormattedElement : BaseFormattedElement
    {
        public string Content { get; set; }

        public TextFormattedElement()
        {
        }

        public TextFormattedElement(string content)
        {
            this.Content = content;
        }

        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append(HttpUtility.HtmlEncode(Content));
        }

        public override OpenXmlElement ToOpenXmlElement()
        {
            return new DocumentFormat.OpenXml.Wordprocessing.Text(Content);
        }
    }
}
