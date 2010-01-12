using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using NHibernate;
using System.Data;

namespace Russell.RADAR.POC.Infrastructure.NH
{
    public class UnitOfWork : IUnitOfWork
    {
        public const string SESSION_CONTEXT_KEY = @"NHSession";
        const string UOW_CONTEXT_KEY = @"UnifOfWork";

        public static IUnitOfWork Current
        {
            get
            {
                return (IUnitOfWork)HttpContext.Current.Items[UOW_CONTEXT_KEY];
            }
        }

        public static ISession CurrentSession
        {
            get { return Current.Session; }
        }

        public static IUnitOfWork Start()
        {
            return Start(IsolationLevel.ReadCommitted);
        }

        public static IUnitOfWork Start(IsolationLevel isolationLevel)
        {
            if (HttpContext.Current.Items[UOW_CONTEXT_KEY] != null)
            {
                throw new ApplicationException(@"Another UnitOfWork has been started in the same context. You should not start multiple UnitOfWorks simultaneously.");
            }

            var result = new UnitOfWork();
            HttpContext.Current.Items[UOW_CONTEXT_KEY] = result;
            
            result.Session = (ISession)HttpContext.Current.Items[SESSION_CONTEXT_KEY];
            result.Session.BeginTransaction(isolationLevel);

            return result;
        }

        public ISession Session { get; private set; }

        private bool hasBeenCommitted = false;

        public void Commit()
        {
            Session.Transaction.Commit();
            hasBeenCommitted = true;
        }

        public void Dispose()
        {
            if (!hasBeenCommitted)
            {
                Session.Transaction.Rollback();
            }
        }
    }
}
