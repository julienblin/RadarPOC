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

        public abstract IEnumerable<OpenXmlElement> ToOpenXmlElements(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDocumentPart);

        private List<IFormattedElement> children = new List<IFormattedElement>();

        public IList<IFormattedElement> Children
        {
            get { return children; }
        }

        public IFormattedElement Parent { get; set; }

        protected void ForEachChild(Action<IFormattedElement> action)
        {
            foreach (var child in Children)
            {
                action(child);
            }
        }

        public abstract object Clone();

        protected void DeepCopyChildren(IEnumerable<IFormattedElement> otherChildren)
        {
            Children.Clear();
            foreach (var child in otherChildren)
            {
                Children.Add((IFormattedElement)child.Clone());
            }
        }
    }
}
