using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSLicenseStatisticsCollection
    {
        internal List<MSLicenseStatistics> _statistics = null;
        internal List<MSLicenseForecast> _forecast = null;
        internal List<MSLicenseStatisticsExt> _history = null;
        public List<MSLicenseStatistics> Statistics => _statistics;
        public List<MSLicenseForecast> Forecast => _forecast;
        public List<MSLicenseStatisticsExt> History => _history;

        public MSLicenseStatisticsCollection()
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            _statistics = db.GetLicenseStatistics();
            _forecast = db.GetLicenseForecast();
            _history = db.GetAllLicenseStatistics();
        }

        public MSLicenseStatisticsCollection(string select, string sku = "")
        {
            string[] elements = { "forecast", "statistics", "history" };

            if (!string.IsNullOrWhiteSpace(select))
            {
                elements = select.Split(',');
                for (int i = 0; i < elements.Length; i++)
                {
                    elements[i] = elements[i].Trim().ToLower();
                }
            }
                
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
         
            if (elements.Contains("statistics"))
                _statistics = string.IsNullOrWhiteSpace(sku) ? db.GetLicenseStatistics() : db.GetLicenseStatistics(sku);

            if (elements.Contains("forecast"))
                _forecast = string.IsNullOrWhiteSpace(sku) ? db.GetLicenseForecast() : db.GetLicenseForecast(sku);

            if (elements.Contains("history"))
                _history = string.IsNullOrWhiteSpace(sku) ? db.GetAllLicenseStatistics() : db.GetAllLicenseStatistics(sku);
        }
    }

    [Route("api/[controller]")]
    [Authorize(Roles = "LicenseReader,Admin")]
    [ApiController]
    public class LicenseStatistics : ControllerBase
    {
        private readonly ILogger<PublicReport> _logger;

        public LicenseStatistics(ILogger<PublicReport> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task<MSLicenseStatisticsCollection> Get(string sku, string select)
        {
            return new MSLicenseStatisticsCollection(select, sku);
        }
    }
}
