using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FluentNHibernate.Mapping;

namespace Russell.RADAR.POC.Entities
{
    public class DocumentMap : ClassMap<Document>
    {
        public DocumentMap()
        {
            Id(x => x.Id);
            Map(x => x.Author);
        }
    }
}
