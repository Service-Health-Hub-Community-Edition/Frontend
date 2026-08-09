using Microsoft.ServiceHealthHub.Core;

namespace Microsoft.ServiceHealthHub
{
    public class M365Roadmap: CommunicationBase
    {
        public M365Roadmap(string id)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            MSRoadmapNotification message = db.GetRoadmapNotification(id);

            if (message != null)
            {
                WorkItemId = message.WorkItemID;
                WorkItemUrl = message.WorkItemURL;
                Message = message;
            }
        }
    }
}
