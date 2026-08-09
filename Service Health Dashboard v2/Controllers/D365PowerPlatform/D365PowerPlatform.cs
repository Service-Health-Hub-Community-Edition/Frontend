using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using System;

namespace Service_Health_Dashboard_v2.Controllers
{
    [ApiController]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    public class D365PowerPlatformRelease : ControllerBase
    {
        private readonly ILogger<AzureUpdate> _logger;

        public D365PowerPlatformRelease(ILogger<AzureUpdate> logger)
        {
            _logger = logger;
        }

        // D365 and Power Platform Releases endpoints
        [Route("api/d365pp/releases")]
        [HttpGet]
        public IActionResult GetD365PowerPlatformReleases()
        {
            var db = new MSSHNotificationDatabase();

            MSUserInfo userInfo = new(User);
            if (!Guid.TryParse(userInfo.ObjectId, out Guid userId)) { userId = Guid.Empty; }

            return Ok(db.GetD365PowerPlatformReleases(userId));
        }

        [Route("api/d365pp/releases/{id}")]
        [HttpGet]
        public IActionResult GetD365PowerPlatformRelease(string id)
        {
            var db = new MSSHNotificationDatabase();

            MSUserInfo userInfo = new(User);
            if (!Guid.TryParse(userInfo.ObjectId, out Guid userId)) { userId = Guid.Empty; }

            if (userId != Guid.Empty)
                db.SetViewpointReadFlag(userId, id, Guid.Empty, Guid.Empty, true);

            return Ok(db.GetD365PowerPlatformRelease(id, userId));
        }
    }
}
