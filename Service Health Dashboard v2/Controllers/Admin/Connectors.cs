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
    public class Connectors : ControllerBase
    {
        private readonly ILogger<Connectors> _logger;

        public Connectors(ILogger<Connectors> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<List<MSConnector>> Get(string type)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            return db.GetConnectors(type);
        }

        [Route("api/[controller]/{id}")]
        [HttpGet]
        public async Task<MSConnector> Get(Guid id)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            return db.GetConnector(id);
        }

        [Route("api/[controller]/definitions/{definitionId}")]
        [HttpGet]
        public async Task<List<MSConnector>> GetAsync(Guid definitionId)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            return db.GetConnectors(definitionId);
        }

        [Route("api/[controller]")]
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] MSConnector body)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.UpdateConnector(body);
            return Ok();
        }

        [Route("api/[controller]")]
        [HttpPatch]
        public async Task<IActionResult> PatchAsync([FromBody] MSConnector body)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.UpdateConnector(body);
            return Ok();
        }

        [Route("api/[controller]")]
        [HttpDelete]
        public async Task<IActionResult> DeleteAsync([FromBody] MSConnector body)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.DeleteConnector(body);
            return Ok();
        }
    }

    [Authorize(Roles = "Admin")]
    [ApiController]
    public class ConnectorDefinition : ControllerBase
    {
        private readonly ILogger<Connectors> _logger;

        public ConnectorDefinition(ILogger<Connectors> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<List<MSConnectorDefinition>> Get(string type)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            return db.GetConnectorDefinitions(type);
        }
    }
}
