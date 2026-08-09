using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSItemArchiveBody
    {
        public string id { get; set; }
        public string type { get; set; }
    }

    public class MSItemArchiveResponse
    {
        public int result { get; set; }
    }

    public class MSWeeklyMessageStat
    {
        public DateTime week { get; set; }
        public int items { get; set; }

        public static List<MSWeeklyMessageStat> GetWeeklyMessageStats(string messageType)
        {
            var db = new MSSHNotificationDatabase();
            var result = db.GetMessageCenterWeeklyStatistics(messageType);
            List<MSWeeklyMessageStat> stats = new();

            foreach (MSPropertyBag item in result)
            {
                stats.Add(new MSWeeklyMessageStat
                {
                    week = item["Week"] is DBNull ? DateTime.MinValue : (DateTime)item["Week"],
                    items = item["Items"] is DBNull ? 0 : (int)item["Items"]
                });
            }

            return stats;
        }
    }

    public class MSCurrentWeekServiceStat
    {
        public string service { get; set; }
        public int items { get; set; }

        public static List<MSCurrentWeekServiceStat> GetServiceStats()
        {
            var db = new MSSHNotificationDatabase();
            var result = db.GetMessageCenterPastWeekStatistics();
            List<MSCurrentWeekServiceStat> stats = new();

            foreach (MSPropertyBag item in result)
            {
                stats.Add(new MSCurrentWeekServiceStat
                {
                    service = item["Service"] is DBNull ? "" : (string)item["Service"],
                    items = item["Items"] is DBNull ? 0 : (int)item["Items"]
                });
            }

            return stats;
        }
    }

    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    [ApiController]
    public class Items : ControllerBase
    {
        private readonly ILogger<Items> _logger;

        public Items(ILogger<Items> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<IActionResult> Get()
        {
            return Ok();
        }

        [Route("api/[controller]/[action]")]
        [HttpPost]
        public async Task<MSItemArchiveResponse> ArchiveAsync([FromBody] MSItemArchiveBody body)
        {
            MSUserInfo userInfo = new MSUserInfo(User);
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            MSItemArchiveResponse res = new MSItemArchiveResponse()
            {
                result = db.SetItemArchiveFlag(body.id, body.type, true, userInfo.UserName)
            };

            return res;
        }

        [Route("api/[controller]/[action]")]
        [HttpPost]
        public async Task<MSItemArchiveResponse> RestoreAsync([FromBody] MSItemArchiveBody body)
        {
            MSUserInfo userInfo = new MSUserInfo(User);
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            MSItemArchiveResponse res = new MSItemArchiveResponse()
            {
                result = db.SetItemArchiveFlag(body.id, body.type, false, userInfo.UserName)
            };

            return res;
        }

        [Route("api/[controller]/[action]")]
        [HttpPost]
        public async Task<IActionResult> ViewpointAsync([FromBody] MSViewpoint body)
        {
            Guid userId;
            MSUserInfo userInfo = new MSUserInfo(User);
            if (!Guid.TryParse(userInfo.ObjectId, out userId)) { userId = Guid.Empty; }

            MSSHNotificationDatabase db = new MSSHNotificationDatabase();

            body.userId = userId;
            if (body.subscriptionId == null)
                body.subscriptionId = Guid.Empty;

            if (body.tenantId == null)
                body.tenantId = Guid.Empty;

            if (body.read != null)
                db.SetViewpointReadFlag(
                    body.userId.Value,
                    body.communicationId,
                    body.tenantId.Value,
                    body.subscriptionId.Value,
                    body.read.Value);

            if (body.archived != null)
                db.SetViewpointArchiveFlag(
                    body.userId.Value,
                    body.communicationId,
                    body.tenantId.Value,
                    body.subscriptionId.Value,
                    body.archived.Value);

            if (body.favorite != null)
                db.SetViewpointFavoriteFlag(
                    body.userId.Value,
                    body.communicationId,
                    body.tenantId.Value,
                    body.subscriptionId.Value,
                    body.favorite.Value);

            return Ok();
        }

        [Route("api/[controller]/statistics/messageCenter")]
        [HttpGet]
        public IActionResult GetMCWeekly()
        {
            return Ok(MSWeeklyMessageStat.GetWeeklyMessageStats("SERVICEUPDATEMESSAGE"));
        }

        [Route("api/[controller]/statistics/serviceHealth")]
        [HttpGet]
        public IActionResult GetHealthIssueWeekly()
        {
            return Ok(MSWeeklyMessageStat.GetWeeklyMessageStats("SERVICEHEALTHISSUE"));
        }

        [Route("api/[controller]/statistics/roadmap")]
        [HttpGet]
        public IActionResult GetRoadmapWeekly()
        {
            return Ok(MSWeeklyMessageStat.GetWeeklyMessageStats("ROADMAPCOMMUNICATION"));
        }

        [Route("api/[controller]/statistics/azureServiceHealth")]
        [HttpGet]
        public IActionResult GetAzureServiceHealthWeekly()
        {
            return Ok(MSWeeklyMessageStat.GetWeeklyMessageStats("AZURESERVICEHEALTHALERT"));
        }

        [Route("api/[controller]/statistics/currentWeek/messageCenter")]
        [HttpGet]
        public IActionResult GetCurrentMCWeeklyStats()
        {
            return Ok(MSCurrentWeekServiceStat.GetServiceStats());
        }
    }
}
