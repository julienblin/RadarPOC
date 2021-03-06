﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using DocumentFormat.OpenXml;

namespace Russell.RADAR.POC.Entities.Content
{
    public class ParagraphFormattedElement : BaseFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<p>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</p>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDocumentPart)
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            ForEachChild(x =>
            {
                if (x is TextFormattedElement)
                {
                    result.Append(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(x.ToOpenXmlElements(mainDocumentPart))
                    );

                }
                else
                {
                    result.Append(x.ToOpenXmlElements(mainDocumentPart));
                }
            });
            return new List<OpenXmlElement> { result };
        }

        public override object Clone()
        {
            var clone = new ParagraphFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
