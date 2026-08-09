using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.ServiceHealthHub.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSServiceHealthEvent
    {
        private string _id = string.Empty;
        private string _title = string.Empty;
        private string _classification = string.Empty;
        private string _workload = string.Empty;
        private string _internalWorkloadName = string.Empty;
        private string _status = string.Empty;
        private bool _published = false;

        public string Id => _id;
        public string Title => _title;
        public string Classification => _classification;
        public string Workload => _workload;
        public string InternalWorkloadName => _internalWorkloadName;
        public string Status => _status;
        public bool Published => _published;

        public MSServiceHealthEvent(string id, string title, string classification, string workload, string internalWorkloadName, string status, bool published = false)
        {
            _id = id;
            _title = title;
            _classification = classification;
            _workload = workload;
            _internalWorkloadName = internalWorkloadName;
            _status = status;
            _published = published;
        }
    }

    public class MSServiceHealthEventCountAnalysis
    {
        private int _advisories30 = 0;
        private int _advisories60 = 0;
        private int _incidents30 = 0;
        private int _incidents60 = 0;
        
        public int Advisories30 => _advisories30;
        public int Advisories60 => _advisories60;
        public int Incidents30 => _incidents30;
        public int Incidents60 => _incidents60;
        
        public MSServiceHealthEventCountAnalysis(
            int advisories30, int advisories60, int incidents30, int incidents60)
        {
            _advisories30 = advisories30;
            _advisories60 = advisories60;
            _incidents30 = incidents30;
            _incidents60 = incidents60;
        }
    }
    public class MSServiceHealthEventsStatistics
    {
        internal MSServiceHealthEventCountAnalysis _pastEvents;
        public MSServiceHealthEventCountAnalysis PastEvents => _pastEvents;

        private MSServiceHealthEventsStatistics()
        {
            
        }

        public static async Task<MSServiceHealthEventsStatistics> Create()
        {
            var stats = new MSServiceHealthEventsStatistics();
            stats.GetPast60Analysis();
            return stats;
        }

        private void GetPast60Analysis()
        {
            List<M365ServiceIssue> comms = Cache.Instance.GetServiceHealthIssueCollection();
            comms = comms?.FindAll(c => c.StartDateTime >= DateTime.UtcNow.AddDays(-60));

            int a30 = 0; int a60 = 0; int i30 = 0; int i60 = 0;

            if (comms != null && comms.Count >= 1)
            {
                var groupedComms = comms
                    .GroupBy(c => c.Classification);

                foreach (var gComm in groupedComms)
                {
                    int past30 = gComm.Where(c => c.StartDateTime > DateTime.UtcNow.AddDays(-30)).Count();
                    int past30to60 = gComm.Where(c => c.StartDateTime < DateTime.UtcNow.AddDays(-30)).Count();

                    string classification = Regex.Replace(gComm.Key.ToString(), "([a-z])([A-Z])", "$1 $2");

                    switch (classification.ToLower().Trim())
                    {
                        case "advisory":
                            a30 = past30;
                            a60 = past30to60;
                            break;
                        case "incident":
                            i30 = past30;
                            i60 = past30to60;
                            break;
                    }
                }

                _pastEvents = new MSServiceHealthEventCountAnalysis(a30, a60, i30, i60);
            }
        }
    }

    public class MSServiceStatistics
    {
        internal MSServiceHealthEventsStatistics _eventStatistics = MSServiceHealthEventsStatistics.Create().Result;

        public MSServiceHealthEventsStatistics EventStatistics => _eventStatistics;
    }

    [Route("api/[controller]")]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    [ApiController]
    public class Statistics : ControllerBase
    {
        private readonly ILogger<Statistics> _logger;

        public Statistics(ILogger<Statistics> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task<MSServiceStatistics> Get()
        {
            return new MSServiceStatistics();
        }
    }
}
