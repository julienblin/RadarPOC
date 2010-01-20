using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Russell.RADAR.POC.Entities.Content
{
    public class TableCellFormattedElement : BaseFormattedElement, IWidthSpecifier
    {
        public System.Web.UI.WebControls.Unit Width { get; set; }

        public int? Rowspan { get; set; }
        public int? Colspan { get; set; }

        public TableCellFormattedElement()
        {
            Width = System.Web.UI.WebControls.Unit.Empty;
            Rowspan = null;
            Colspan = null;
        }

        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<td");
            if (!Width.IsEmpty)
                builder.AppendFormat(" style=\"width: {0}\"", Width);

            if (Rowspan.HasValue)
                builder.AppendFormat(" rowspan=\"{0}\"", Rowspan);

            if (Colspan.HasValue)
                builder.AppendFormat(" colspan=\"{0}\"", Colspan);

            builder.Append(">");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</td>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            TableCell result = new TableCell();
            var cellProperties = new TableCellProperties();

            if (!Width.IsEmpty)
            {
                var cellWidth = UnitHelper.Convert(Width).To<TableCellWidth>();
                cellProperties.Append(cellWidth);
            }

            if (Colspan.HasValue)
            {
                var gridSpan = new GridSpan() { Val = Colspan };
                cellProperties.Append(gridSpan);
            }

            if (Rowspan.HasValue)
            {
                var verticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                cellProperties.Append(verticalMerge);
            }

            result.Append(cellProperties);

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
            clone.Width = Width;
            clone.Rowspan = Rowspan;
            clone.Colspan = Colspan;
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
