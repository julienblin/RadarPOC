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
using Russell.RADAR.POC.Entities;

namespace Russell.RADAR.POC.WebApp.OpinionDocuments.Components
{
    public partial class SectionEditor : System.Web.UI.UserControl
    {
        public string Title
        {
            get { return labelSection.Text; }
            set { labelSection.Text = value; }
        }

        public OpinionDocumentSection Section { get; set; }

        public int GetRank()
        {
            return ddlRank.SelectedIndex + 1;
        }

        public string GetContent()
        {
            return textBoxSection.Text;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (!Page.IsPostBack)
            {
                ddlRank.DataSource = new int[] { 1, 2, 3, 4, 5 };
                ddlRank.DataBind();

                if (Section != null)
                {
                    textBoxSection.Text = Section.Content.ToXHTML();
                    ddlRank.SelectedIndex = Section.Rank - 1;
                }
            }
        }
    }
}