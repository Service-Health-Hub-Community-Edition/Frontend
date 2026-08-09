using Microsoft.ServiceHealthHub.Core;
using Microsoft.ServiceHealthHub.Graph;
using System.Text.RegularExpressions;

namespace Microsoft.ServiceHealthHub
{
    public class M365ServiceHealthIssue: CommunicationBase
    {
        public M365ServiceHealthIssue(string id)
        {
            var issue = GraphApiClientHelper.Client.Admin.ServiceAnnouncement.Issues[id].GetAsync().Result;

            issue.AdditionalData.Add("status", Regex.Replace(issue.Status.ToString(), "([a-z])([A-Z])", "$1 $2"));
            issue.AdditionalData.Add("classification", Regex.Replace(issue.Classification.ToString(), "([a-z])([A-Z])", "$1 $2"));

            MSServiceNotification svcNotificationInfo = MSServiceNotification.GetNotification(id);
            if (!string.IsNullOrWhiteSpace(svcNotificationInfo.WorkItemID))
            {
                issue.AdditionalData.Add("workItemId", svcNotificationInfo.WorkItemID);
                issue.AdditionalData.Add("workItemUrl", svcNotificationInfo.WorkItemURL);

                WorkItemId = svcNotificationInfo.WorkItemID;
                WorkItemUrl = svcNotificationInfo.WorkItemURL;
            }

            issue.AdditionalData.Add("public", svcNotificationInfo.Public);
            issue.AdditionalData.Add("comments", svcNotificationInfo.Comments);

            Message = issue;
        }
    }
}
