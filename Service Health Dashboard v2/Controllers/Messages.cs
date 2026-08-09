using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.OData.Query;
using Microsoft.AspNetCore.SignalR.Protocol;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.ServiceHealthHub.Graph;
using Microsoft.VisualStudio.Services.Notification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Service_Health_Dashboard_v2.Controllers
{
    [ApiController]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    public class Messages : ControllerBase
    {
        private readonly ILogger<Messages> _logger;

        public Messages(ILogger<Messages> logger)
        {
            _logger = logger;
        }

        private void GetViewpoint(List<M365MessageCenterItem> res)
        {
            Guid userId;
            MSUserInfo userInfo = new MSUserInfo(User);
            if (!Guid.TryParse(userInfo.ObjectId, out userId)) { userId = Guid.Empty; }

            List<MSViewpoint> viewpoints = new List<MSViewpoint>();

            if (userId != Guid.Empty && res != null && res.Count > 0)
            {
                MSCommunicationIdCollection idCollection = new MSCommunicationIdCollection();

                foreach (M365MessageCenterItem message in res)
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
                foreach (M365MessageCenterItem message in res)
                {
                    MSViewpoint viewpoint = viewpoints.Find(v => v.communicationId == message.Id); // add check for tenantId and subscriptionId
                    message.ServiceHealthHubViewpoint.archived = viewpoint != null && viewpoint.archived.Value;
                    message.ServiceHealthHubViewpoint.favorited = viewpoint != null && viewpoint.favorite.Value;
                    message.ServiceHealthHubViewpoint.viewed = viewpoint != null && viewpoint.read.Value;
                }
            }
        }

        private void GetTags(List<M365MessageCenterItem> res)
        {
            Guid userId;
            MSUserInfo userInfo = new MSUserInfo(User);
            if (!Guid.TryParse(userInfo.ObjectId, out userId)) { userId = Guid.Empty; }

            List<MSCommunicationTag> tags = new List<MSCommunicationTag>();

            if (userId != Guid.Empty && res != null && res.Count > 0)
            {
                MSCommunicationIdCollection idCollection = new MSCommunicationIdCollection();

                foreach (M365MessageCenterItem message in res)
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
                foreach (M365MessageCenterItem message in res)
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
            List<M365MessageCenterItem> res = Cache.Instance.GetMessageCenterCollection();

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
            M365MessageCenterItem message = Cache.Instance.GetMessageCenterItem(id);
            MSSHNotificationDatabase db = new();           
            Guid userId;
            MSUserInfo userInfo = new MSUserInfo(User);
            if (!Guid.TryParse(userInfo.ObjectId, out userId)) { userId = Guid.Empty; }

            if (message == null)
            {
                List<string> idList = new List<string>();
                List<MSServiceNotification> dbNotifications = new List<MSServiceNotification>();
                idList.Add(id);

                db.GetNotifications(idList, dbNotifications);

                if (dbNotifications.Count > 0)
                {
                    ServiceUpdateMessage mc = Newtonsoft.Json.JsonConvert.DeserializeObject<ServiceUpdateMessage>(dbNotifications[0].Data);

                    message = M365MessageCenterItem.From(mc);

                    if (message.AdditionalData == null)
                        message.AdditionalData = new Dictionary<string, object>();

                    List<MSSummaryCacheEntry> summaryCache = db.GetLastNotificationSummary(idList);

                    message.Public = dbNotifications[0].Public;
                    message.PublishingComments = dbNotifications[0].Comments;
                    message.ExtendedProperties = dbNotifications[0].ExtendedProperties;
                    message.Summary = summaryCache[0];
                    message.ServiceHealthHubState = dbNotifications[0].State;

                    if (!string.IsNullOrWhiteSpace(dbNotifications[0].WorkItemID))
                    {
                        message.Task.TaskId = dbNotifications[0].WorkItemID;
                        message.Task.TaskUrl = dbNotifications[0].WorkItemURL;
                    }
                }
                
            }

            if (message != null)
            {
                List<M365MessageCenterItem> items = new() { message };

                if (userId != Guid.Empty)
                    db.SetViewpointReadFlag(userId, message.Id, Guid.Empty, Guid.Empty, true);

                GetViewpoint(items);
                GetTags(items);
            }

            if (message != null)
                return Ok(message);
            else
                return NotFound();
        }
    }
}
