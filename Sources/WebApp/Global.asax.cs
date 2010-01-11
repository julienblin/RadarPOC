using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.SessionState;
using Castle.Windsor;

namespace Russell.RADAR.POC.WebApp
{
    public class Global : System.Web.HttpApplication, IContainerAccessor
    {
        private static WebAppContainer container;

        public IWindsorContainer Container
        {
            get { return container; }
        }

        protected void Application_Start(object sender, EventArgs e)
        {
            container = new WebAppContainer();
        }

        protected void Application_End(object sender, EventArgs e)
        {
            container.Dispose();
        }
    }
}