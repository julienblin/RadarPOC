using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;

namespace Russell.RADAR.POC.PublishingServices
{
    public class InDesignPublishingService : IPublishingService
    {
        public byte[] PublishAsPDF(Document document)
        {
            InDesign.Application app = (InDesign.Application)COMCreateObject("InDesign.Application");

            return new byte[] { };
        }

        private static object COMCreateObject(string sProgID)
        {
            // We get the type using just the ProgID
            Type oType = Type.GetTypeFromProgID(sProgID);
            if (oType != null)
            {
                return Activator.CreateInstance(oType);
            }

            return null;
        }
    }
}
