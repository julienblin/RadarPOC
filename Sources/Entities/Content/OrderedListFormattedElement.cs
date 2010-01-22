using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace Russell.RADAR.POC.Entities.Content
{
    public class OrderedListFormattedElement : ListFormattedElement
    {
        public int NumberingInstanceId { get; private set; }

        public override void ToXHTML(StringBuilder builder)
        {
            builder.Append("<ol>");
            ForEachChild(x => x.ToXHTML(builder));
            builder.Append("</ol>");
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDocumentPart)
        {
            // Creates a new numbered instance for each Ordered list.
            NumberingInstanceId = IdHelper.GenerateIntId();

            var numberingDefinition = FindOrCreateNumberingDefinitionPart(mainDocumentPart);
            var numbering = new Numbering();
            numberingDefinition.Numbering = numbering;

            NumberingInstance numberingInstance = new NumberingInstance() { NumberID = NumberingInstanceId };
            numbering.Append(numberingInstance);

            AbstractNumId abstractNumId = new AbstractNumId() { Val = 0 };
            numberingInstance.Append(abstractNumId);

            var result = new List<OpenXmlElement>();
            ForEachChild(x =>
            {
                Debug.Assert(x is ListItemFormattedElement);
                result.AddRange(x.ToOpenXmlElements(mainDocumentPart));
            });
            return result;
        }

        private NumberingDefinitionsPart FindOrCreateNumberingDefinitionPart(MainDocumentPart mainDocumentPart)
        {
            NumberingDefinitionsPart result = null;

            var existingDefinitions = mainDocumentPart.GetPartsOfType<NumberingDefinitionsPart>();

            foreach (NumberingDefinitionsPart existingDefinition in existingDefinitions)
	        {
        		 result = existingDefinition;
	        }

            if (result == null)
                result = mainDocumentPart.AddNewPart<NumberingDefinitionsPart>();

            return result;
        }

        public override object Clone()
        {
            var clone = new OrderedListFormattedElement();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
