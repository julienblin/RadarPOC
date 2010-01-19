using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace Russell.RADAR.POC.Entities.Content
{
    public class UnorderedListFormattedContent : ListFormattedContent
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<ul>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</ul>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            var result = new List<OpenXmlElement>();
            ForEachChild(x =>
            {
                Debug.Assert(x is ListItemFormattedContent);
                result.AddRange(x.ToOpenXmlElements());
            });
            return result;
        }

        public override object Clone()
        {
            var clone = new UnorderedListFormattedContent();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
