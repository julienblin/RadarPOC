using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.Entities.Content
{
    public class ItalicFormattedContent : BaseFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<em>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</em>");
        }

        public override DocumentFormat.OpenXml.OpenXmlElement ToOpenXmlElement()
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Run();
            var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            var italicProperty = new DocumentFormat.OpenXml.Wordprocessing.Italic();

            runProperties.Append(italicProperty);
            result.Append(runProperties);

            ForEachChild(x => result.Append(x.ToOpenXmlElement()));
            return result;
        }
    }
}
