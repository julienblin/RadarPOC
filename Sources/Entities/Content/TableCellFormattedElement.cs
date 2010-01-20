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

        public TableCellFormattedElement()
        {
            Width = System.Web.UI.WebControls.Unit.Empty;
        }

        public override void ToXHTML(StringBuilder builder)
        {
            if (Width.IsEmpty)
            {
                builder.Append("<td>");
            }
            else
            {
                builder.AppendFormat("<td style=\"width: {0}\">", Width);
            }
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</td>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            TableCell result = new TableCell();

            if (!Width.IsEmpty)
            {
                var cellProperties = new TableCellProperties();
                var cellWidth = UnitHelper.Convert(Width).To<TableCellWidth>();
                cellProperties.Append(cellWidth);
                result.Append(cellProperties);
            }

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
