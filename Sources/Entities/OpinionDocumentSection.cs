using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.Entities
{
    public class OpinionDocumentSection
    {
        int rank = 1;
        public virtual int Rank
        {
            get { return rank;  }
            set { rank = value;  }
        }

        public virtual string Content { get; set; }
    }
}
