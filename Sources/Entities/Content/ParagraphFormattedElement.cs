using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace Russell.RADAR.POC.Entities.Content
{
    public class ParagraphFormattedElement : BaseFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<p>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</p>");
        }

        public override DocumentFormat.OpenXml.OpenXmlElement ToOpenXmlElement()
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            ForEachChild(x => 
                result.Append(x.ToOpenXmlElement())
            );
            return result;
        }

        public override object Clone()
        {
            var clone = new ParagraphFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
