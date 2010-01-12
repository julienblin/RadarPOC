using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using Castle.Windsor;

namespace Russell.RADAR.POC.WebApp
{
    public class BasePage : System.Web.UI.Page
    {
        public IContainerAccessor ContainerAccessor
        {
            get { return (IContainerAccessor)HttpContext.Current.ApplicationInstance; }
        }

        public T Resolve<T>()
        {
            return ContainerAccessor.Container.Resolve<T>();
        }
    }
}
