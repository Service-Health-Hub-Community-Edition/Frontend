using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Authorization;
using Microsoft.ServiceHealthHub.Core;
using System.Collections.Generic;
using System;
using Microsoft.Graph;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSJobHistory
    {
        public Guid correlationId { get; set; }
        public string jobName { get; set; }
        public DateTime? start { get; set; }
        public DateTime? end { get; set; }
        public string state { get; set; }
        public int itemsCreated { get; set; }
        public int itemsModified { get; set; }
        public int itemsFailed { get; set; }
        public int tasksCreated { get; set; }
        public int tasksModified { get; set; }
        public int tasksFailed { get; set; }
        public int notificationsSent { get; set; }
        public int notificationsFailed { get; set; }

        public MSJobHistory() {
        
        }
        
        public MSJobHistory(MSPropertyBag dbRecord) {
            correlationId = dbRecord["CorrelationId"] is DBNull ? Guid.Empty : (Guid)dbRecord["CorrelationId"];
            jobName = dbRecord["JobName"] is DBNull ? string.Empty : (string)dbRecord["JobName"];
            start = dbRecord["Start"] is DBNull ? null : (DateTime)dbRecord["Start"];
            end = dbRecord["End"] is DBNull ? null : (DateTime)dbRecord["End"];
            itemsCreated = dbRecord["ItemsCreated"] is DBNull ? 0 : (int)dbRecord["ItemsCreated"];
            itemsModified = dbRecord["ItemsModified"] is DBNull ? 0 : (int)dbRecord["ItemsModified"];
            itemsFailed = dbRecord["ItemsFailed"] is DBNull ? 0 : (int)dbRecord["ItemsFailed"];
            tasksCreated = dbRecord["TasksCreated"] is DBNull ? 0 : (int)dbRecord["TasksCreated"];
            tasksModified = dbRecord["TasksModified"] is DBNull ? 0 : (int)dbRecord["TasksModified"];
            tasksFailed = dbRecord["TasksFailed"] is DBNull ? 0 : (int)dbRecord["TasksFailed"];
            notificationsSent = dbRecord["NotificationsSent"] is DBNull ? 0 : (int)dbRecord["NotificationsSent"];
            notificationsFailed = dbRecord["NotificationsFailed"] is DBNull ? 0 : (int)dbRecord["NotificationsFailed"];

            string jobState = dbRecord["JobState"] is DBNull ? string.Empty : (string)dbRecord["JobState"];
            if (string.IsNullOrWhiteSpace(jobState))
                jobState = "Running";

            switch (jobState)
            {
                case "Running":
                    MSSHNotificationDatabase db = new MSSHNotificationDatabase();
                    bool timedOut = db.GetJobTimeoutState(correlationId);
                    if (timedOut)
                        state = "Timed out";
                    else
                        state = jobState;
                    break;
                case "JobCompleted":
                    state = itemsFailed > 0 || tasksFailed > 0 || notificationsFailed > 0 ? "Failed" : "Completed";
                    break;
                case "JobFailed":
                    state = "Failed";
                    break;
                default:
                    state = itemsFailed > 0 || tasksFailed > 0 || notificationsFailed > 0 ? "Failed" : "Completed";
                    break;
                
            }
        }
    }

    [Authorize(Roles = "Admin")]
    [ApiController]
    public class JobHistory : ControllerBase
    {
        private readonly ILogger<JobHistory> _logger;

        public JobHistory(ILogger<JobHistory> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("api/jobs/{id}/history")]
        public IActionResult Get(Guid id, int? top, string? communicationId, bool? hideEmptyJobs, DateTime? startDate, DateTime? endDate)
        {
            MSSHNotificationDatabase db = new();
            List<MSPropertyBag> results = new();

            if (!string.IsNullOrWhiteSpace(communicationId))
                results = db.GetJobStatistics(id, communicationId);
            else
                results = db.GetJobStatistics(id, 
                    top == null ? 0 : top.Value, 
                    hideEmptyJobs == null ? null : hideEmptyJobs.Value,
                    startDate == null ? null : startDate.Value,
                    endDate == null ? null : endDate.Value);
            
            List<MSJobHistory>history = new();
            foreach (MSPropertyBag result in results)
                history.Add(new MSJobHistory(result));

            return Ok(history);
        }
    }
}
