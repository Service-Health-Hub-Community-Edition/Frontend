using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.OData.Query;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Microsoft.ServiceHealthHub.Core;
using System.Collections.Generic;
using System.Linq;

namespace Service_Health_Dashboard_v2.Controllers
{
    [ApiController]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    [Route("api/microsoft365/healthOverview")]
    public class ServiceStatus : ControllerBase
    {
        private readonly ILogger<ServiceStatus> _logger;

        public ServiceStatus(ILogger<ServiceStatus> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [EnableQuery]
        public IActionResult Get()
        {
            List<ServiceHealth> res = Cache.Instance.GetServiceHealthOverviewCollection();

            if (res != null)
                return Ok(res.AsQueryable());
            else
                return NotFound();
        }
    }
}
