using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Russell.RADAR.POC.Entities;

namespace Russell.RADAR.POC.PublishingServices
{
    public interface IPublishingService
    {
        byte[] Publish(Document document);
    }
}
