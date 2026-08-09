using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Authorization;
using Microsoft.ServiceHealthHub.Core;
using System.Collections.Generic;
using System;
using System.Threading;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSJob
    {
        public string icon { get; set; }
        public Guid id { get; set; }
        public string name { get; set; }
        public DateTime? lastRun { get; set; }
        public string? duration { get; set; }
        public string state { get; set; }
    }

    [Authorize(Roles = "Admin")]
    [ApiController]
    public class Jobs : ControllerBase
    {
        private readonly ILogger<Jobs> _logger;

        public Jobs(ILogger<Jobs> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("api/jobs")]
        public IActionResult Get()
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            List<MSComponent> components = db.GetComponent(component: null);
            List<MSJob> jobs = new List<MSJob>();

            foreach (var component in components)
            {
                if (component.Capabilities != null &&
                    (component.Capabilities.Contains("Sync") || component.Capabilities.Contains("Webhook")))
                {
                    List<MSPropertyBag> runtimeInfo = db.GetJobStatistics(component.Id, 1, false, null, null);
                    MSJobHistory jobHistory = runtimeInfo.Count > 0 ? new MSJobHistory(runtimeInfo[0]) : null;
                    jobs.Add(jobHistory != null ?
                        new MSJob
                        {
                            icon = component.Icon,
                            id = component.Id,
                            name = component.Name,
                            lastRun = jobHistory.start,
                            duration = string.Format(@"{0:hh\:mm\:ss\.ff}", (jobHistory.end - jobHistory.start)),
                            state = jobHistory.state
                        } :
                        new MSJob {
                            icon = component.Icon,
                            name = component.Name,
                            id = component.Id }
                        );
                }
            }

            return Ok(jobs);
        }
    }
}
