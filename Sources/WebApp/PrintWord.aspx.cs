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
using System.IO;

namespace Russell.RADAR.POC.WebApp
{
    public partial class PrintWord : BasePage
    {
        public Document document;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (document == null)
            {
                using (IUnitOfWork uow = UnitOfWork.Start())
                {
                    var authoringService = Resolve<IAuthoringService>();
                    document = authoringService.Retrieve(Convert.ToInt32(Request.Params["id"]));
                }
            }

            Response.ContentType = @"application/msword";
            Response.AddHeader("Content-Disposition", "attachment; filename=GeneratedDocument.docx");

            using (var memStream = new MemoryStream())
            {
                document.StreamOpenXMLDocument(memStream);
                Response.BinaryWrite(memStream.ToArray());
            }

            Response.Flush();
            Response.End();
        }
    }
}
