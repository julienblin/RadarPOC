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
using Russell.RADAR.POC.PublishingServices;
using System.Diagnostics;

namespace Russell.RADAR.POC.WebApp.OpinionDocuments
{
    public partial class PrintPDF : BasePage
    {
        public OpinionDocument document;

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

            Response.ContentType = "application/pdf";
            var publishingService = Resolve<IPublishingService>();

            switch (publishingService.ExportOption)
            {
                case ExportOption.AsByte:
                    Response.BinaryWrite(publishingService.PublishAsPDF(document));
                    break;
                case ExportOption.AsFile:
                    Response.WriteFile(publishingService.PublishAsPDFFile(document));
                    break;
                default:
                    Debug.Assert(false);
                    break;
            }

            Response.Flush();
        }
    }
}
