using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.Entities
{
    public abstract class Document
    {
        public virtual int Id { get; set; }
        public virtual string Author { get; set; }
    }
}
