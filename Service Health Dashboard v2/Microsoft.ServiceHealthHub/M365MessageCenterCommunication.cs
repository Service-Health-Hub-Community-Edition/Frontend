using Microsoft.ServiceHealthHub.Graph;
using Microsoft.ServiceHealthHub.Core;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Microsoft.ServiceHealthHub
{
    public class M365MessageCenterCommunication: CommunicationBase
    {
        public M365MessageCenterCommunication(string id)
        {
            var message = GraphApiClientHelper.Client.Admin.ServiceAnnouncement.Messages[id].GetAsync().Result;

            TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
            message.AdditionalData.Add("category", myTI.ToTitleCase(Regex.Replace(message.Category.ToString(), "([a-z])([A-Z])", "$1 $2")));

            MSServiceNotification svcNotificationInfo = MSServiceNotification.GetNotification(id);
            if (!string.IsNullOrWhiteSpace(svcNotificationInfo.WorkItemID))
            {
                message.AdditionalData.Add("workItemId", svcNotificationInfo.WorkItemID);
                message.AdditionalData.Add("workItemUrl", svcNotificationInfo.WorkItemURL);

                WorkItemId = svcNotificationInfo.WorkItemID;
                WorkItemUrl = svcNotificationInfo.WorkItemURL;
            }

            message.AdditionalData.Add("public", svcNotificationInfo.Public);
            message.AdditionalData.Add("comments", svcNotificationInfo.Comments);

            Message = message;
        }
    }
}
