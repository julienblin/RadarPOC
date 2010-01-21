using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using System.Diagnostics;

namespace Russell.RADAR.POC.Entities.Content
{
    public class OrderedListFormattedElement : ListFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<ol>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</ol>");
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
            var clone = new OrderedListFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
