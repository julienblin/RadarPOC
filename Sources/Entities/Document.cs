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

        public virtual DocumentState State { get; set; }

        public abstract DocumentType DocumentType { get; }

        public virtual bool CanBeEdited()
        {
            return (State != DocumentState.Published);
        }
    }
}
