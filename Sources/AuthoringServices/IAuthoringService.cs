using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;

namespace Russell.RADAR.POC.AuthoringServices
{
    public interface IAuthoringService
    {
        Document Create(ContentType contentType);
        Document Retrieve(int id);
        void Save(Document document);
    }
}
