using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.OData.Query;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Microsoft.ServiceHealthHub;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.ServiceHealthHub.Graph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Service_Health_Dashboard_v2.Controllers
{
    [ApiController]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    public class Issues : ControllerBase
    {
        private readonly ILogger<Issues> _logger;

        public Issues(ILogger<Issues> logger)
        {
            _logger = logger;
        }

        private void GetViewpoint(List<M365ServiceIssue> res)
        {
            Guid userId;
            MSUserInfo userInfo = new MSUserInfo(User);
            if (!Guid.TryParse(userInfo.ObjectId, out userId)) { userId = Guid.Empty; }

            List<MSViewpoint> viewpoints = new List<MSViewpoint>();

            if (userId != Guid.Empty && res != null && res.Count > 0)
            {
                MSCommunicationIdCollection idCollection = new MSCommunicationIdCollection();

                foreach (M365ServiceIssue message in res)
                {
                    idCollection.Add(new MSCommunicationId()
                    {
                        userId = userId,
                        communicationId = message.Id,
                        tenantId = Guid.Empty,        // tenantId and subscriptionId shall be adjusted
                        subscriptionId = Guid.Empty   // during multitenancy implementation
                    });
                }

                MSSHNotificationDatabase db = new MSSHNotificationDatabase();
                viewpoints = db.GetViewpoints(idCollection);
            }

            lock (res)
            {
                foreach (M365ServiceIssue message in res)
                {
                    MSViewpoint viewpoint = viewpoints.Find(v => v.communicationId == message.Id); // add check for tenantId and subscriptionId
                    message.ServiceHealthHubViewpoint.archived = viewpoint != null && viewpoint.archived.Value;
                    message.ServiceHealthHubViewpoint.favorited = viewpoint != null && viewpoint.favorite.Value;
                    message.ServiceHealthHubViewpoint.viewed = viewpoint != null && viewpoint.read.Value;
                }
            }
        }

        private void GetTags(List<M365ServiceIssue> res)
        {
            Guid userId;
            MSUserInfo userInfo = new MSUserInfo(User);
            if (!Guid.TryParse(userInfo.ObjectId, out userId)) { userId = Guid.Empty; }

            List<MSCommunicationTag> tags = new List<MSCommunicationTag>();

            if (userId != Guid.Empty && res != null && res.Count > 0)
            {
                MSCommunicationIdCollection idCollection = new MSCommunicationIdCollection();

                foreach (M365ServiceIssue message in res)
                {
                    idCollection.Add(new MSCommunicationId()
                    {
                        userId = userId,
                        communicationId = message.Id,
                        tenantId = Guid.Empty,        // tenantId and subscriptionId shall be adjusted
                        subscriptionId = Guid.Empty   // during multitenancy implementation
                    });
                }

                MSSHNotificationDatabase db = new MSSHNotificationDatabase();
                tags = db.GetTagCollection(idCollection);
            }

            lock (res)
            {
                foreach (M365ServiceIssue message in res)
                {
                    List<MSCommunicationTag> commTags = tags.FindAll(t => t.MessageId == message.Id && t.TagId != null);
                    message.OrganizationTags = commTags;
                }
            }
        }

        [HttpGet]
        [EnableQuery]
        [Route("api/[controller]")]
        public IActionResult Get()
        {
            List<M365ServiceIssue> res = Cache.Instance.GetServiceHealthIssueCollection();

            GetViewpoint(res);
            GetTags(res);

            if (res != null)
                return Ok(res.AsQueryable());
            else
                return NotFound();
        }

        [HttpGet]
        [Route("api/[controller]/{id?}")]
        public IActionResult Get(string id)
        {
            M365ServiceIssue issue = Cache.Instance.GetServiceHealthIssue(id);

            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            Guid userId;
            MSUserInfo userInfo = new MSUserInfo(User);
            if (!Guid.TryParse(userInfo.ObjectId, out userId)) { userId = Guid.Empty; }

            if (issue == null)
            {
                List<string> idList = new List<string>();
                List<MSServiceNotification> dbNotifications = new List<MSServiceNotification>();
                idList.Add(id);

                db.GetNotifications(idList, dbNotifications);

                if (dbNotifications.Count > 0)
                {
                    ServiceHealthIssue i = Newtonsoft.Json.JsonConvert.DeserializeObject<ServiceHealthIssue>(dbNotifications[0].Data);

                    issue = M365ServiceIssue.From(i);

                    if (issue.AdditionalData == null)
                        issue.AdditionalData = new Dictionary<string, object>();

                    List<MSSummaryCacheEntry> summaryCache = db.GetLastNotificationSummary(idList);

                    issue.Public = dbNotifications[0].Public;
                    issue.PublishingComments = dbNotifications[0].Comments;
                    issue.ExtendedProperties = dbNotifications[0].ExtendedProperties;
                    issue.Summary = summaryCache[0];
                    issue.ServiceHealthHubState = dbNotifications[0].State;

                    if (!string.IsNullOrWhiteSpace(dbNotifications[0].WorkItemID))
                    {
                        issue.Task.TaskId = dbNotifications[0].WorkItemID;
                        issue.Task.TaskUrl = dbNotifications[0].WorkItemURL;
                    }
                }
            }

            if (issue != null)
            {
                List<M365ServiceIssue> items = new() { issue };

                if (userId != Guid.Empty)
                    db.SetViewpointReadFlag(userId, issue.Id, Guid.Empty, Guid.Empty, true);

                GetViewpoint(items);
                GetTags(items);
            }

            if (issue != null)
                return Ok(issue);
            else
                return NotFound();
        }
    }
}
