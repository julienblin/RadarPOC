using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.Entities.Content
{
    public class BoldFormattedElement : BaseFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<strong>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</strong>");
        }

        public override DocumentFormat.OpenXml.OpenXmlElement ToOpenXmlElement()
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Run();
            var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            var boldProperty = new DocumentFormat.OpenXml.Wordprocessing.Bold();

            runProperties.Append(boldProperty);
            result.Append(runProperties);

            ForEachChild(x => result.Append(x.ToOpenXmlElement()));
            return result;
        }

        public override object Clone()
        {
            var clone = new BoldFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
