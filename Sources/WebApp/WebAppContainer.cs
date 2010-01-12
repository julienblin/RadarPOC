﻿using System;
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

namespace Russell.RADAR.POC.WebApp
{
    public class WebAppContainer : WindsorContainer
    {
        public WebAppContainer()
        {
            Register(
                Component.For<ISessionFactory>()
                    .UsingFactoryMethod(
                        () => ConfigureFluentNHibernate()
                    )
            );
        }

        private ISessionFactory ConfigureFluentNHibernate()
        {
            return Fluently.Configure()
                .Database(SQLiteConfiguration.Standard.ConnectionString(@"Radar"))
                .Mappings(m =>
                    m.FluentMappings.AddFromAssemblyOf<Document>())
                .BuildSessionFactory();
        }
    }
}
