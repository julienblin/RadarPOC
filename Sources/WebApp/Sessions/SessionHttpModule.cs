using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NHibernate;
using Castle.Windsor;
using Russell.RADAR.POC.Infrastructure.NH;

namespace Russell.RADAR.POC.WebApp.Sessions
{
    public class SessionHttpModule : IHttpModule
    {
        private IContainerAccessor containerAccessor;

        public void Init(HttpApplication context)
        {
            containerAccessor = (IContainerAccessor)context;

            context.BeginRequest += new EventHandler(OpenSession);
            context.EndRequest += new EventHandler(CloseSession);
        }
 
        private void OpenSession(object sender, EventArgs e)
        {
            var sessionFactory = containerAccessor.Container.Resolve<ISessionFactory>();
            var newSession = sessionFactory.OpenSession();
            HttpContext.Current.Items[UnitOfWork.SESSION_CONTEXT_KEY] = newSession;
        }
 
        private void CloseSession(object sender, EventArgs e)
        {
            var currentSession = (ISession)HttpContext.Current.Items[UnitOfWork.SESSION_CONTEXT_KEY];
            currentSession.Flush();
            currentSession.Close();
        }

        public void Dispose()
        {
            HttpContext.Current.Items[UnitOfWork.SESSION_CONTEXT_KEY] = null;
        }
    }
}
