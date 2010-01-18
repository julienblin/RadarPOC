using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace Russell.RADAR.POC.Entities.Content
{
    public interface IFormattedElement
    {
        void ToXHTML(StringBuilder builder);
        OpenXmlElement ToOpenXmlElement();

        IList<IFormattedElement> Children { get; }
    }
}
