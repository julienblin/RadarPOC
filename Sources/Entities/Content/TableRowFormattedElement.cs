using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace Russell.RADAR.POC.Entities.Content
{
    public class TableRowFormattedElement : BaseFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<tr>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</tr>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDocumentPart)
        {
            TableRow result = new TableRow();
            ForEachChild(x =>
            {
                Debug.Assert(x is TableCellFormattedElement);
                result.Append(x.ToOpenXmlElements(mainDocumentPart));
            });
            return new List<OpenXmlElement> { result };
        }

        public override object Clone()
        {
            var clone = new TableRowFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
