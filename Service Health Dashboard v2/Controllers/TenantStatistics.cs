using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.AspNetCore.Authorization;

namespace Service_Health_Dashboard_v2.Controllers
{
    [Route("api/[controller]")]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Public,Admin,LicenseReader")]
    [ApiController]
    public class MonthlyActiveUsers : ControllerBase
    {
        private readonly ILogger<MonthlyActiveUsers> _logger;

        public MonthlyActiveUsers(ILogger<MonthlyActiveUsers> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task<List<MSMonthlyActiveUsers>> Get()
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            List<MSMonthlyActiveUsers> mauStatistics = db.GetMAU();
            return mauStatistics;
        }
    }
}
