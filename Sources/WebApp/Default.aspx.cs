using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Russell.RADAR.POC.Infrastructure.NH;
using Russell.RADAR.POC.WebApp;
using Russell.RADAR.POC.AuthoringServices;
using Russell.RADAR.POC.Entities;

namespace WebApp
{
    public partial class _Default : BasePage
    {
        protected override void OnInitComplete(EventArgs e)
        {
            base.OnInitComplete(e);
            linkNewOpinionDocument.Click += new EventHandler(linkNewOpinionDocument_Click);
        }

        protected void Page_Load(object sender, EventArgs e)
        {

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

            Response.Redirect(string.Format("~/OpinionDocuments/Edit.aspx?id={0}", newDoc.Id));
        }
    }
}
