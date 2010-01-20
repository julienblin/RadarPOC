using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using HtmlAgilityPack;
using System.Web;
using System.Text.RegularExpressions;

namespace Russell.RADAR.POC.Entities.Content
{
    public class FormattedContent : BaseFormattedElement
    {
        static readonly Regex reStyleAttr = new Regex(@"(?<name>[^;:]+)\s*:\s*(?<value>[^;:]+)", RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public void FromXHTML(string input)
        {
            Children.Clear();

            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(input);

            ParseRecursive(this, doc.DocumentNode);
        }

        private void ParseRecursive(IFormattedElement baseElement, HtmlNode baseNode)
        {
            foreach (var childNode in baseNode.ChildNodes)
            {
                IFormattedElement createdNode = null;
                switch (childNode.Name)
                {
                    case "p":
                        createdNode = new ParagraphFormattedElement();
                        break;
                    case "b":
                    case "strong":
                        createdNode = new BoldFormattedElement();
                        break;
                    case "i":
                    case "em":
                        createdNode = new ItalicFormattedElement();
                        break;
                    case "ul":
                        createdNode = new UnorderedListFormattedElement();
                        break;
                    case "ol":
                        createdNode = new OrderedListFormattedElement();
                        break;
                    case "li":
                        createdNode = new ListItemFormattedElement();
                        break;
                    case "table":
                        createdNode = new TableFormattedElement();
                        ExtractStyleWidth(childNode, (IWidthSpecifier)createdNode);
                        break;
                    case "tbody":
                    case "thead":
                        // ignore
                        break;
                    case "tr":
                        createdNode = new TableRowFormattedElement();
                        break;
                    case "td":
                    case "th":
                        createdNode = new TableCellFormattedElement();
                        break;
                    default:
                        createdNode = new TextFormattedElement(TrimText(HttpUtility.HtmlDecode(childNode.InnerText)));
                        break;
                }
                if (createdNode != null)
                {
                    baseElement.Children.Add(createdNode);
                    createdNode.Parent = baseElement;
                    ParseRecursive(createdNode, childNode);
                }
                else
                {
                    // if element was skipped (e.g. tbody)
                    ParseRecursive(baseElement, childNode);
                }
            }
        }

        private void ExtractStyleWidth(HtmlNode htmlNode, IWidthSpecifier content)
        {
            if (htmlNode.Attributes.Contains("style"))
            {
                var attrMatches = reStyleAttr.Matches(htmlNode.Attributes["style"].Value);
                foreach (Match attrMatch in attrMatches)
                {
                    if (attrMatch.Groups["name"].Value.Equals("width", StringComparison.InvariantCultureIgnoreCase))
                    {
                        content.Width = System.Web.UI.WebControls.Unit.Parse(attrMatch.Groups["value"].Value);
                    }
                }
            }
        }

        /// <summary>
        /// Removes non-sgnificant characters (\r\n\t) in xhtml string.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string TrimText(string text)
        {
            return Regex.Replace(text, "[\r\n\t]", string.Empty);
        }

        public string ToXHTML()
        {
            var builder = new StringBuilder();
            ToXHTML(builder);
            return builder.ToString();
        }

        public override void ToXHTML(StringBuilder builder)
        {
            ForEachChild(x => x.ToXHTML(builder));
        }

        public override IEnumerable<OpenXmlElement> ToOpenXmlElements()
        {
            var result = new List<OpenXmlElement>();
            ForEachChild(x =>
            {
                if (x is ParagraphFormattedElement)
                {
                    result.AddRange(x.ToOpenXmlElements());
                }

            });
            return result;
        }

        public override object Clone()
        {
            var clone = new FormattedContent();
            clone.DeepCopyChildren(Children);
            return clone;
        }
    }
}
