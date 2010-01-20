using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace Russell.RADAR.POC.Entities.Content
{
    public class TableFormattedElement : ParagraphFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<table>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</table>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            Table result = new Table();
            ForEachChild(x =>
            {
                Debug.Assert(x is TableRowFormattedElement);
                result.Append(x.ToOpenXmlElements());
            });
            return new List<OpenXmlElement> { result };
        }

        public override object Clone()
        {
            var clone = new TableFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
