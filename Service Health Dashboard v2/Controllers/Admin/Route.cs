using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.AspNetCore.Authorization;

namespace Service_Health_Dashboard_v2.Controllers
{
    [Authorize(Roles = "Admin")]
    [ApiController]
    public class Route : ControllerBase
    {
        private readonly ILogger<Connectors> _logger;

        public Route(ILogger<Connectors> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<List<MSRoute>> Get(Guid? component, string type)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            return db.GetRoute(component, type);
        }

        [Route("api/[controller]")]
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] MSRouteInt body)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.UpdateRoute(body);
            return Ok();
        }

        [Route("api/[controller]")]
        [HttpPatch]
        public async Task<IActionResult> PatchAsync([FromBody] MSRouteInt body)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.UpdateRoute(body);
            return Ok();
        }

        [Route("api/[controller]")]
        [HttpDelete]
        public async Task<IActionResult> DeleteAsync([FromBody] MSRouteId body)
        {
            MSRouteInt routeToDelete = body.GetRouteObject();
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.DeleteRoute(routeToDelete);
            return Ok();
        }
    }
}
