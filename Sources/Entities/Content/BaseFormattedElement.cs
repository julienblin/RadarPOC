using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace Russell.RADAR.POC.Entities.Content
{
    public abstract class BaseFormattedElement : IFormattedElement
    {
        public abstract void ToXHTML(StringBuilder builder);

        public abstract OpenXmlElement ToOpenXmlElement();

        private List<IFormattedElement> children = new List<IFormattedElement>();

        public IList<IFormattedElement> Children
        {
            get { return children; }
        }

        protected void ForEachChild(Action<IFormattedElement> action)
        {
            foreach (var child in Children)
            {
                action(child);
            }
        }
    }
}
