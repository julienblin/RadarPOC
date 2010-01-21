using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Russell.RADAR.POC.Entities.Content
{
    public interface IFormattedElement : ICloneable
    {
        void ToXHTML(StringBuilder builder);
        IEnumerable<OpenXmlElement> ToOpenXmlElements(MainDocumentPart mainDocumentPart);

        IList<IFormattedElement> Children { get; }

        IFormattedElement Parent { get; set; }
    }
}
