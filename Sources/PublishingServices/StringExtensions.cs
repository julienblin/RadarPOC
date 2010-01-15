using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.PublishingServices
{
    internal static class StringExtensions
    {
        /// <summary>
        /// Creates a full xhtml document (including html and body tags) from an xhtml fragment.
        /// Basically wraps a string between html and body tags.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToCompleteXHTML(this string input)
        {
            return string.Format("<html><head/><body><div style=\"font-family: Arial; font-size: 11pt;\">{0}</div></body></html>", input);
        }
    }
}
