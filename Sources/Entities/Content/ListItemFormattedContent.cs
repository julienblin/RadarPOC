using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Russell.RADAR.POC.Entities.Content
{
    public class ListItemFormattedContent : BaseFormattedElement
    {
        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<li>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</li>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();

            var paragraphProperties = new ParagraphProperties();
            var paragraphStyleId = new ParagraphStyleId() { Val = "UnorderedListStyle" };

            var numberingProperties = new NumberingProperties();
            var numberingLevelReference = new NumberingLevelReference() { Val = 0 };

            NumberingId numberingId;
            if (Parent is UnorderedListFormattedContent)
            {
                numberingId = new NumberingId() { Val = 3 };
            }
            else
            {
                numberingId = new NumberingId() { Val = 2 };
            }


            numberingProperties.Append(numberingLevelReference);
            numberingProperties.Append(numberingId);

            var spacingBetweenLines= new SpacingBetweenLines() { After = "100", AfterAutoSpacing = true };

            paragraphProperties.Append(paragraphStyleId);
            paragraphProperties.Append(numberingProperties);
            paragraphProperties.Append(spacingBetweenLines);
            result.Append(paragraphProperties);

            ForEachChild(x =>
            {
                if (x is TextFormattedElement)
                {
                    result.Append(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(x.ToOpenXmlElements())
                    );

                }
                else
                {
                    result.Append(x.ToOpenXmlElements());
                }
            });

            return new List<OpenXmlElement> { result };
        }

        public override object Clone()
        {
            var clone = new ListItemFormattedContent();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
