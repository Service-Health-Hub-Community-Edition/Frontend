using Microsoft.ServiceHealthHub.Core;
using System.Data.SqlTypes;
using System;
using Microsoft.VisualStudio.Services.Notification;
using Newtonsoft.Json;

namespace Microsoft.ServiceHealthHub
{
    public class MSAzureUpdate
    {
        public string id { get; set; }
        public string title { get; set; }
        public string releaseStatus { get; set; }
        public DateTime published { get; set; }
        public object tags { get; set; }
        public string summary { get; set; }
        public string link { get; set; }
        public object contents { get; set; }
    }

    public class AzureUpdate: CommunicationBase
    {
        public AzureUpdate(string id)
        {
            /* MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            // MSAzureUpdate message = new MSAzureUpdate(id);

            if (message != null)
            {
                WorkItemId = message.WorkItemID;
                WorkItemUrl = message.WorkItemURL;
                Message = message;
            } */
        }
    }
}
