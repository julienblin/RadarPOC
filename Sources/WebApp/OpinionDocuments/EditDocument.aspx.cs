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
                //textAreaDiscussion.Value = doc.Discussion;
            }
        }
    }
}
