using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace Russell.RADAR.POC.Entities.Content
{
    public class TableFormattedElement : ParagraphFormattedElement, IWidthSpecifier
    {
        public System.Web.UI.WebControls.Unit Width { get; set; }

        public TableFormattedElement()
        {
            Width = System.Web.UI.WebControls.Unit.Percentage(100);
        }

        public override void ToXHTML(StringBuilder builder)
        {
            builder.AppendFormat("<table border=\"1\" cellpadding=\"0\" cellspacing=\"0\" style=\"width: {0}\">", Width);
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</table>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            Table result = new Table();

            var tableProperties = new TableProperties();
            var tableStyle = new TableStyle() { Val = "GrilledutableauSimpleTable" };
            var tableWidth = UnitHelper.Convert(Width).To<TableWidth>();
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
            clone.Width = Width;
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
