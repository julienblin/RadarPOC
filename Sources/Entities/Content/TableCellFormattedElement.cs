using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Russell.RADAR.POC.Entities.Content
{
    public class TableCellFormattedElement : BaseFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<td>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</td>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            TableCell result = new TableCell();
            var paraContent = new Paragraph();
            ForEachChild(x =>
            {
                if (x is TextFormattedElement)
                {
                    paraContent.Append(
                        new Run(x.ToOpenXmlElements())
                    );

                }
                else
                {
                    paraContent.Append(x.ToOpenXmlElements());
                }
            });
            result.Append(paraContent);
            return new List<OpenXmlElement> { result };
        }

        public override object Clone()
        {
            var clone = new TableCellFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
