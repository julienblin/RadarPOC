using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NHibernate;
using Castle.MicroKernel;

namespace Russell.RADAR.POC.WebApp.Sessions
{
    public class SessionHttpModule : IHttpModule
    {
        public void Init(HttpApplication context)
        {
            context.BeginRequest += new EventHandler(OpenSession);
            context.EndRequest += new EventHandler(CloseSession);
        }
 
        private void OpenSession(object sender, EventArgs e)
        {
            
        }
 
        private void CloseSession(object sender, EventArgs e)
        {
            
        }

        public void Dispose()
        {
        }
    }
}
