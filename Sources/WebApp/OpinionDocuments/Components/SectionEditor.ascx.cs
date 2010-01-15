using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

namespace Russell.RADAR.POC.WebApp.OpinionDocuments.Components
{
    public partial class SectionEditor : System.Web.UI.UserControl
    {
        public Label Label
        {
            get { return labelSection; }
        }

        public TextBox TextBox
        {
            get { return textBoxSection; }
        }

        public string Title
        {
            get { return labelSection.Text; }
            set { labelSection.Text = value; }
        }
    }
}