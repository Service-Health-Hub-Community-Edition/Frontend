using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSActiveUsage
    {
        public string service { get; set; }
        public int active { get; set; }
        public int inactive { get; set; }
        public DateTime reportDate { get; set; }

        public static List<MSActiveUsage> GetMonthlyActiveReport(string[]? services)
        {
            var db = new MSSHNotificationDatabase();
            var result = db.GetMonthlyActiveUsersReport(services);
            List<MSActiveUsage> stats = new();

            foreach (MSPropertyBag item in result)
            {
                if (item["Service"] is not DBNull &&
                    item["ReportDate"] is not DBNull)
                stats.Add(new MSActiveUsage
                {
                    service = (string)item["Service"],
                    reportDate = (DateTime)item["ReportDate"],
                    active = item["Active"] is DBNull ? 0 : (int)item["Active"],
                    inactive = item["Inactive"] is DBNull ? 0 : (int)item["Inactive"]
                });
            }

            return stats;
        }
    }

    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    [ApiController]
    public class MonthlyActiveUserReport : ControllerBase
    {
        private readonly ILogger<MonthlyActiveUserReport> _logger;

        public MonthlyActiveUserReport(ILogger<MonthlyActiveUserReport> logger)
        {
            _logger = logger;
        }

        [Route("api/reports/monthlyActiveUsers")]
        [HttpGet]
        public async Task<IActionResult> Get(string? service)
        {
            if (string.IsNullOrWhiteSpace(service))
                return Ok(MSActiveUsage.GetMonthlyActiveReport(null));
            else
            {
                string[] serviceList = service.Split(',');
                return Ok(MSActiveUsage.GetMonthlyActiveReport(serviceList));
            }
        }
    }
}
