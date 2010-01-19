using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Castle.Windsor;
using Castle.MicroKernel.Registration;
using NHibernate;
using FluentNHibernate.Cfg;
using FluentNHibernate.Cfg.Db;
using Russell.RADAR.POC.Entities;
using System.Configuration;
using Russell.RADAR.POC.AuthoringServices;
using Castle.Facilities.FactorySupport;
using Castle.Windsor.Configuration.Interpreters;

namespace Russell.RADAR.POC.WebApp
{
    public class WebAppContainer : WindsorContainer
    {
        public WebAppContainer()
        {
            RegisterFacilities();
            RegisterNHibernate();
            RegisterServices();
        }

        private void RegisterFacilities()
        {
            AddFacility<FactorySupportFacility>();
        }

        private void RegisterNHibernate()
        {
            Register(
                Component.For<ISessionFactory>()
                    .UsingFactoryMethod(
                        () => ConfigureFluentNHibernate()
                    )
            );
        }

        private void RegisterServices()
        {
            Register(
                Component.For<IAuthoringService>().ImplementedBy<NHAuthoringService>()
            );
        }

        private ISessionFactory ConfigureFluentNHibernate()
        {
            switch (ConfigurationManager.AppSettings["DbType"])
            {
                case "sqlite":
                    return Fluently.Configure()
                        .Database(SQLiteConfiguration.Standard.ConnectionString(x => x.FromConnectionStringWithKey(@"Radar")))
                        .Mappings(m =>
                            m.FluentMappings.AddFromAssemblyOf<Document>())
                        .BuildSessionFactory();
                case "sqlserver":
                    return Fluently.Configure()
                        .Database(MsSqlConfiguration.MsSql2000.ConnectionString(x => x.FromConnectionStringWithKey(@"Radar")))
                        .Mappings(m =>
                            m.FluentMappings.AddFromAssemblyOf<Document>())
                        .BuildSessionFactory();
                default:
                    throw new NotSupportedException("Unknow dbtype: " + ConfigurationManager.AppSettings["DbType"]);
            }
        }
    }
}
