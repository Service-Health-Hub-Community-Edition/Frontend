using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    [Authorize(Roles = "Admin")]
    [ApiController]
    public class Components : ControllerBase
    {
        private readonly ILogger<Connectors> _logger;

        public Components(ILogger<Connectors> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<List<MSComponent>> Get(Guid? id)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            return db.GetComponent(id);
        }
    }
}
