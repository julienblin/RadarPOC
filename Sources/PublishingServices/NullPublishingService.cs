using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.PublishingServices
{
    public class NullPublishingService : IPublishingService
    {
        #region IPublishingService Members

        public byte[] Publish(Russell.RADAR.POC.Entities.Document document)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
