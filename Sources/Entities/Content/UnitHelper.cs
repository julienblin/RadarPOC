using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace Russell.RADAR.POC.Entities.Content
{
    public static class UnitHelper
    {
        const int WORD_DOCUMENT_RESOLUTION_IN_DPI = 72;
        const double PAPER_USLETTER_WIDTH_IN_INCHES = 8.3;
        const double PAPER_USLETTER_WIDTH_IN_POINTS = WORD_DOCUMENT_RESOLUTION_IN_DPI * PAPER_USLETTER_WIDTH_IN_INCHES;
        const double PAPER_USLETTER_WIDTH_IN_DXA = PAPER_USLETTER_WIDTH_IN_POINTS * 20;
        const int XHTML_EDITING_TARGET_WIDTH_IN_PIXELS = 800;
        const double PIXEL_IN_DXA_CONVERSION = PAPER_USLETTER_WIDTH_IN_DXA / XHTML_EDITING_TARGET_WIDTH_IN_PIXELS;

        public static WordUnit Convert(System.Web.UI.WebControls.Unit unit)
        {
            var result = new WordUnit();

            switch (unit.Type)
            {
                case System.Web.UI.WebControls.UnitType.Percentage:
                    result.Width = (System.Convert.ToInt32(unit.Value) * 50).ToString();
                    result.Type = TableWidthUnitValues.Pct;
                    break;
                case System.Web.UI.WebControls.UnitType.Pixel:
                    result.Width = (System.Convert.ToInt32(unit.Value) * PIXEL_IN_DXA_CONVERSION).ToString();
                    result.Type = TableWidthUnitValues.Dxa;
                    break;
                default:
                    Debug.Assert(false, "Unsupported width type : " + unit.Type.ToString());
                    break;
            }

            return result;
        }

        public static long ConvertPixelsToEMUS(System.Web.UI.WebControls.Unit unit)
        {
            if (unit.Type != System.Web.UI.WebControls.UnitType.Pixel)
                throw new NotSupportedException();

            return System.Convert.ToInt64(unit.Value * 9525);
        }
    }

    public class WordUnit
    {
        public EnumValue<TableWidthUnitValues> Type { get; set; }
        public StringValue Width { get; set; }

        public T To<T>()
            where T : TableWidthType, new()
        {
            return new T { Width = Width, Type = Type };
        }
    }
}
