using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FluentNHibernate.Mapping;

namespace Russell.RADAR.POC.Entities
{
    public class OpinionDocumentMap : SubclassMap<OpinionDocument>
    {
        public OpinionDocumentMap()
        {
            Map(x => x.Discussion);
            Map(x => x.InvestmentStaff);
        }
    }
}
