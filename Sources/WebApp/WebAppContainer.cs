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
using Russell.RADAR.POC.PublishingServices;

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
                Component.For<IAuthoringService>().ImplementedBy<NHAuthoringService>(),
                Component.For<IPublishingService>().ImplementedBy<TelerikReportPublishingService>()
            );
        }

        private ISessionFactory ConfigureFluentNHibernate()
        {
            return Fluently.Configure()
                .Database(SQLiteConfiguration.Standard.ConnectionString(x => x.FromConnectionStringWithKey(@"Radar")))
                .Mappings(m =>
                    m.FluentMappings.AddFromAssemblyOf<Document>())
                .BuildSessionFactory();
        }
    }
}
