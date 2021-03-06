﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace Russell.RADAR.POC.Entities.Content
{
    public class ItalicFormattedElement : BaseFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<em>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</em>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDocumentPart)
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Run();
            var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            var italicProperty = new DocumentFormat.OpenXml.Wordprocessing.Italic();

            runProperties.Append(italicProperty);
            result.Append(runProperties);

            ForEachChild(x => result.Append(x.ToOpenXmlElements(mainDocumentPart)));
            return new List<OpenXmlElement> { result };
        }

        public override object Clone()
        {
            var clone = new ItalicFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
