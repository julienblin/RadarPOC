using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;
using NHibernate;
using Russell.RADAR.POC.Infrastructure.NH;

namespace Russell.RADAR.POC.AuthoringServices
{
    public class NHAuthoringService : IAuthoringService
    {
        public Document Create(DocumentType contentType)
        {
            Document newDoc = null;
            switch (contentType)
            {
                case DocumentType.Opinion:
                    newDoc = new OpinionDocument(); 
                    break;
                default:
                    throw new ApplicationException(string.Format("Unknown contentType {0}", contentType));
            }

            newDoc.State = DocumentState.Draft;

            UnitOfWork.CurrentSession.Save(newDoc);
            return newDoc;
        }

        public Document Retrieve(int id)
        {
            throw new NotImplementedException();
        }

        public void Save(Document document)
        {
            throw new NotImplementedException();
        }
    }
}
