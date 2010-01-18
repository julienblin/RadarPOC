using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Russell.RADAR.POC.ContentServices
{
    public interface IContentService
    {
        string StripXHTML(string input);
        Paragraph CreateWordParagraphFromXHTML(string input);
    }
}
