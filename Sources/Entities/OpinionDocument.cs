using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.Entities
{
    public class OpinionDocument : Document
    {
        public virtual string Discussion { get; set; }
        public virtual string InvestmentStaff { get; set; }
    }
}
