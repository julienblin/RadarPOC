using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;

namespace Russell.RADAR.POC.AuthoringServices
{
    public interface IAuthoringService
    {
        Document Create(DocumentType contentType);
        Document Retrieve(int id);

        IEnumerable<Document> ListAll();

        void Delete(Document document);
    }
}
