using Microsoft.Graph.Models;
using Microsoft.ServiceHealthHub.Core;
using System.Text.RegularExpressions;

namespace Microsoft.ServiceHealthHub.Graph
{
    public class MSServiceHealthHubTaskInfo
    {
        public string TaskId { get; set; }
        public string TaskUrl { get; set; }
    }

    public class MSServiceHealthHubViewpoint
    {
        private bool m_viewed = false;
        private bool m_archived { get; set; }
        private bool m_favorited { get; set; }

        public bool viewed { get { return m_viewed; } set { m_viewed = value; } }
        public bool archived { get { return m_archived; } set { m_archived = value; } }
        public bool favorited { get { return m_favorited; } set { m_favorited = value; } }
    }

    public class M365ServiceIssue : ServiceHealthIssue
    {
        public List<string> AffectedServices { get; set; }
        public string StatusDisplayName { get; set; }
        public string OriginDisplayName { get; set; }
        public string ClassificationDisplayName { get; set; }
        public bool Public { get; set; }
        public string PublishingComments { get; set; }
        public dynamic ExtendedProperties { get; set; }
        public MSSummaryCacheEntry? Summary { get; set; }
        public MSServiceHealthHubTaskInfo Task { get; set; }
        public string? ServiceHealthHubState { get; set; }
        public MSServiceHealthHubViewpoint ServiceHealthHubViewpoint { get; set; }
        public List<MSCommunicationTag> OrganizationTags { get; set; } = new List<MSCommunicationTag>();

        public static M365ServiceIssue From(ServiceHealthIssue issue)
        {
            M365ServiceIssue result = new M365ServiceIssue()
            {
                Id = issue.Id,
                AdditionalData = issue.AdditionalData,
                Classification = issue.Classification,
                Details = issue.Details,
                EndDateTime = issue.EndDateTime,
                Feature = issue.Feature,
                FeatureGroup = issue.FeatureGroup,
                ImpactDescription = issue.ImpactDescription,
                IsResolved = issue.IsResolved,
                LastModifiedDateTime = issue.LastModifiedDateTime,
                Origin = issue.Origin,
                Posts = issue.Posts,
                Service = issue.Service,
                StartDateTime = issue.StartDateTime,
                Status = issue.Status,
                Title = issue.Title,            
                AffectedServices = new(),
                Task = new(),
                ServiceHealthHubViewpoint = new()
            };

            result.AffectedServices.Add(result.Service);
            if (result.Details != null)
            {
                var childWorkloads = result.Details.FirstOrDefault(i => i.Name == "AffectedChildWorkloads");
                if (childWorkloads != null)
                {
                    var childWorkloadArray = childWorkloads.Value.Split(',');
                    foreach (string workload in childWorkloadArray)
                        result.AffectedServices.Add(workload.Trim());
                }
            }

            return result;
        }
    }

    public class M365MessageCenterItem : ServiceUpdateMessage
    {
        public string CategoryDisplayName { get; set; }
        public bool Public { get; set; }
        public string PublishingComments { get; set; }
        public dynamic ExtendedProperties { get; set; }
        public MSSummaryCacheEntry? Summary { get; set; }
        public MSServiceHealthHubTaskInfo Task { get; set; }
        public string? ServiceHealthHubState { get; set; }
        public MSServiceHealthHubViewpoint ServiceHealthHubViewpoint { get; set; }
        public List<MSCommunicationTag> OrganizationTags { get; set; } = new List<MSCommunicationTag>();

        private static string UppercaseFirst(string s)
        {
            // Check for empty string.
            if (string.IsNullOrEmpty(s))
            {
                return string.Empty;
            }
            // Return char and concat substring.
            return char.ToUpper(s[0]) + s.Substring(1);
        }

        public static M365MessageCenterItem From(ServiceUpdateMessage issue)
        {
            M365MessageCenterItem result = new M365MessageCenterItem()
            {
                Id = issue.Id,
                AdditionalData = issue.AdditionalData,
                Details = issue.Details,
                EndDateTime = issue.EndDateTime,
                ActionRequiredByDateTime = issue.ActionRequiredByDateTime,
                Category = issue.Category,
                CategoryDisplayName = UppercaseFirst(Regex.Replace(issue.Category.ToString(), "([a-z])([A-Z])", "$1 $2").ToLower()),
                IsMajorChange = issue.IsMajorChange,
                Severity = issue.Severity,
                Tags = issue.Tags,
                ViewPoint = issue.ViewPoint,
                LastModifiedDateTime = issue.LastModifiedDateTime,
                Body = issue.Body,
                Services = issue.Services,
                StartDateTime = issue.StartDateTime,
                Title = issue.Title,
                Task = new(),
                ServiceHealthHubViewpoint = new()
            };

            return result;
        }
    }
}
