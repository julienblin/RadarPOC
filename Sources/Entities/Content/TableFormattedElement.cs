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
            builder.AppendFormat("<table style=\"width: {0}\">", Width);
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</table>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            Table result = new Table();

            var tableProperties = new TableProperties();
            var tableStyle = new TableStyle() { Val = "GrilledutableauSimpleTable" };
            var tableWidth = ConvertWidth();
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

        const int WORD_DOCUMENT_RESOLUTION_IN_DPI = 72;
        const double PAPER_USLETTER_WIDTH_IN_INCHES = 8.3;
        const double PAPER_USLETTER_WIDTH_IN_POINTS = WORD_DOCUMENT_RESOLUTION_IN_DPI * PAPER_USLETTER_WIDTH_IN_INCHES;
        const double PAPER_USLETTER_WIDTH_IN_DXA = PAPER_USLETTER_WIDTH_IN_POINTS * 20;
        const int XHTML_EDITING_TARGET_WIDTH_IN_PIXELS = 800;
        const double PIXEL_IN_DXA_CONVERSION = PAPER_USLETTER_WIDTH_IN_DXA / XHTML_EDITING_TARGET_WIDTH_IN_PIXELS;

        private TableWidth ConvertWidth()
        {
            var result = new TableWidth();

            switch (Width.Type)
            {
                case System.Web.UI.WebControls.UnitType.Percentage:
                    result.Width = (Convert.ToInt32(Width.Value) * 50).ToString();
                    result.Type = TableWidthUnitValues.Pct;
                    break;
                case System.Web.UI.WebControls.UnitType.Pixel:
                    result.Width = (Convert.ToInt32(Width.Value) * PIXEL_IN_DXA_CONVERSION).ToString();
                    result.Type = TableWidthUnitValues.Dxa;
                    break;
                default:
                    Debug.Assert(false, "Unsupported width type : " + Width.Type.ToString());
                    break;
            }

            return result;
        }

        public override object Clone()
        {
            var clone = new TableFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
