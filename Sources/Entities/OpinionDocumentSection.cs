using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities.Content;

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

        FormattedContent content = new FormattedContent();
        public virtual FormattedContent Content
        {
            get { return content; }
            set { content = value; }
        }
    }
}
