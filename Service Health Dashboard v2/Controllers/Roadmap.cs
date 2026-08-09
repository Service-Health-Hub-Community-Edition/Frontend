using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    [ApiController]
    public class Roadmap : ControllerBase
    {
        private readonly ILogger<Roadmap> _logger;

        public Roadmap(ILogger<Roadmap> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("api/[controller]")]
        public async Task<MSRoadmapNotificationCollection> Get()
        {
            MSRoadmapNotificationCollection notifications = MSRoadmapNotification.GetNotifications();
            return notifications;
        }

        [HttpGet]
        [Route("api/[controller]/{id?}")]
        public async Task<MSRoadmapNotification> Get(string id)
        {
            MSRoadmapNotification notification = MSRoadmapNotification.GetNotification(id);
            return notification;
        }
    }
}
