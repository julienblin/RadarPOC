using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Xml.Linq;
using Russell.RADAR.POC.Entities;
using Russell.RADAR.POC.Infrastructure.NH;
using Russell.RADAR.POC.AuthoringServices;

namespace Russell.RADAR.POC.WebApp.OpinionDocuments
{
    public partial class EditDocument : BasePage
    {
        OpinionDocument doc;

        protected override void OnInitComplete(EventArgs e)
        {
            base.OnInitComplete(e);
            buttonOk.Click += new EventHandler(buttonOk_Click);
            buttonOK2.Click += new EventHandler(buttonOk_Click);
            buttonCancel.Click += new EventHandler(buttonCancel_Click);
            buttonCancel2.Click += new EventHandler(buttonCancel_Click);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            using (var uow = UnitOfWork.Start())
            {
                var authoringService = Resolve<IAuthoringService>();
                doc = (OpinionDocument)authoringService.Retrieve(Convert.ToInt32(Request.Params["id"]));
                uow.Commit();
            }

            if (!IsPostBack)
            {
                ddlOverallRank.DataSource = new int[] { 1, 2, 3, 4 };
                ddlOverallRank.DataBind();

                textBoxOverall.Text = doc.OverallEvaluationContent;
                ddlOverallRank.SelectedIndex = doc.OverallEvaluationRank - 1;

                textBoxDiscussion.Text = doc.Discussion.ToXHTML();

                SectionInvestementStaff.Section = doc.InvestmentStaff;
                SectionOrganizationalStability.Section = doc.OrganizationalStability;
                SectionAssetAllocation.Section = doc.AssetAllocation;
            }
        }

        void buttonOk_Click(object sender, EventArgs e)
        {
            using (var uow = UnitOfWork.Start())
            {
                doc.OverallEvaluationContent = textBoxOverall.Text;
                doc.OverallEvaluationRank = ddlOverallRank.SelectedIndex + 1;

                doc.Discussion.FromXHTML(textBoxDiscussion.Text);

                doc.InvestmentStaff.Rank = SectionInvestementStaff.GetRank();
                doc.InvestmentStaff.Content.FromXHTML(SectionInvestementStaff.GetContent());
                doc.OrganizationalStability.Rank = SectionOrganizationalStability.GetRank();
                doc.OrganizationalStability.Content.FromXHTML(SectionOrganizationalStability.GetContent());
                doc.AssetAllocation.Rank = SectionAssetAllocation.GetRank();
                doc.AssetAllocation.Content.FromXHTML(SectionAssetAllocation.GetContent());
                uow.Commit();
            }
            Response.Redirect("~/Default.aspx");
        }

        void buttonCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Default.aspx");
        }
    }
}
