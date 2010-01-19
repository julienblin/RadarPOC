using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Russell.RADAR.POC.Entities
{
    public abstract class Document
    {
        public virtual int Id { get; set; }
        public virtual string Author { get; set; }

        public abstract DocumentType DocumentType { get; }

        public abstract void StreamOpenXMLDocument(Stream stream);
    }
}
