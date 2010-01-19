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
using Russell.RADAR.POC.Infrastructure.NH;
using Russell.RADAR.POC.AuthoringServices;
using Russell.RADAR.POC.Entities;

namespace Russell.RADAR.POC.WebApp
{
    public partial class Default : BasePage
    {
        protected override void OnInitComplete(EventArgs e)
        {
            base.OnInitComplete(e);
            linkNewOpinionDocument.Click += new ImageClickEventHandler(linkNewOpinionDocument_Click);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            LoadDocuments();
        }

        private void LoadDocuments()
        {
            using (IUnitOfWork uow = UnitOfWork.Start())
            {
                var authoringService = Resolve<IAuthoringService>();
                repeaterDocuments.DataSource = authoringService.ListAll();
                repeaterDocuments.DataBind();
                uow.Commit();
            }
        }

        void linkNewOpinionDocument_Click(object sender, EventArgs e)
        {
            Document newDoc = null;

            using (IUnitOfWork uow = UnitOfWork.Start())
            {
                var authoringService = Resolve<IAuthoringService>();
                newDoc = authoringService.Create(DocumentType.Opinion);
                uow.Commit();
            }

            Response.Redirect(string.Format("~/OpinionDocuments/EditDocument.aspx?id={0}", newDoc.Id));
        }

        public void Delete_Click(object sender, EventArgs e)
        {
            using (IUnitOfWork uow = UnitOfWork.Start())
            {
                var authoringService = Resolve<IAuthoringService>();
                Document doc = authoringService.Retrieve(Convert.ToInt32(((LinkButton)sender).CommandArgument));
                authoringService.Delete(doc);
                uow.Commit();
            }
            LoadDocuments();
        }
    }
}
