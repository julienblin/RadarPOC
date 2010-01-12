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
using Russell.RADAR.POC.Infrastructure.NH;
using Russell.RADAR.POC.AuthoringServices;

namespace Russell.RADAR.POC.WebApp.OpinionDocuments
{
    public partial class Edit : BasePage
    {
        public OpinionDocument document;

        protected override void OnInitComplete(EventArgs e)
        {
            base.OnInitComplete(e);
            buttonSave.Click += new EventHandler(buttonSave_Click);
            buttonCancel.Click += new EventHandler(buttonCancel_Click);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (document == null)
            {
                using (IUnitOfWork uow = UnitOfWork.Start())
                {
                    var authoringService = Resolve<IAuthoringService>();
                    document = (OpinionDocument)authoringService.Retrieve(Convert.ToInt32(Request.Params["id"]));
                    uow.Commit();
                }
            }

            if (!IsPostBack)
            {
                editorDiscussion.Content = document.Discussion;
                editorInvestmentStaff.Content = document.InvestmentStaff;
            }
        }

        void buttonSave_Click(object sender, EventArgs e)
        {
            using (IUnitOfWork uow = UnitOfWork.Start())
            {
                document.Discussion = editorDiscussion.Content;
                document.InvestmentStaff = editorInvestmentStaff.Content;
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
