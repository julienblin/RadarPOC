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

            var tableProperties = new TableProperties();
            var tableStyle = new TableStyle() { Val = "GrilledutableauSimpleTable" };
            var tableWidth = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            var tableIndentation = new TableIndentation() { Width = 534, Type = TableWidthUnitValues.Dxa };
            var tableLook = new TableLook() { Val = "04A0" };

            tableProperties.Append(tableStyle);
            tableProperties.Append(tableWidth);
            tableProperties.Append(tableIndentation);
            tableProperties.Append(tableLook);

            result.Append(tableProperties);

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
