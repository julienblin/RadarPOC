using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace Russell.RADAR.POC.Entities.Content
{
    public class UnorderedListFormattedElement : ListFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<ul>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</ul>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDocumentPart)
        {
            var result = new List<OpenXmlElement>();
            ForEachChild(x =>
            {
                Debug.Assert(x is ListItemFormattedElement);
                result.AddRange(x.ToOpenXmlElements(mainDocumentPart));
            });
            return result;
        }

        public override object Clone()
        {
            var clone = new UnorderedListFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
