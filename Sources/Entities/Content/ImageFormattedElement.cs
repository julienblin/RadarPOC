using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace Russell.RADAR.POC.Entities.Content
{
    public class ImageFormattedElement : BaseFormattedElement, IWidthSpecifier, IHeightSpecifier
    {
        public string Source { get; set; }

        public System.Web.UI.WebControls.Unit Width { get; set; }

        public System.Web.UI.WebControls.Unit Height { get; set; }

        public ImageFormattedElement()
        {
            Width = System.Web.UI.WebControls.Unit.Empty;
            Height = System.Web.UI.WebControls.Unit.Empty;
        }

        public override void ToXHTML(StringBuilder builder)
        {
            builder.AppendFormat("<img src=\"{0}\" / style=\"", Source);

            if (!Width.IsEmpty)
                builder.AppendFormat(" width: {0};", Width);

            if (!Height.IsEmpty)
                builder.AppendFormat(" height: {0};", Height);

            builder.AppendFormat("\">", Source);
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            throw new NotImplementedException();
        }

        public override object Clone()
        {
            var clone = new ImageFormattedElement();
            clone.Source = Source;
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
