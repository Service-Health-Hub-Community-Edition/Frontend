using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.Process.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;

namespace Service_Health_Dashboard_v2.Controllers
{
    [ApiController]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    public class AzureServiceHealth : ControllerBase
    {
        private readonly ILogger<AzureServiceHealth> _logger;

        public AzureServiceHealth(ILogger<AzureServiceHealth> logger)
        {
            _logger = logger;
        }

        // Azure Health Alerts endpoints
        [Route("api/azure/health/alerts")]
        [HttpGet]
        public IActionResult GetAzureServiceHealthAlerts(bool? active)
        {
            var db = new MSSHNotificationDatabase();

            MSUserInfo userInfo = new(User);
            if (!Guid.TryParse(userInfo.ObjectId, out Guid userId)) { userId = Guid.Empty; }

            return Ok(db.GetAzureHealthAlerts(active, userId));
        }

        [Route("api/azure/health/alerts/{id}")]
        [HttpGet]
        public IActionResult GetAzureServiceHealthAlert(string id)
        {
            var db = new MSSHNotificationDatabase();
            
            MSUserInfo userInfo = new(User);
            if (!Guid.TryParse(userInfo.ObjectId, out Guid userId)) { userId = Guid.Empty; }

            if (userId != Guid.Empty)
                db.SetViewpointReadFlag(userId, id, Guid.Empty, Guid.Empty, true);

            return Ok(db.GetAzureHealthAlert(id, userId));
        }

        [Route("api/azure/health/statistics")]
        [HttpGet]
        public IActionResult GetAzureServiceHealthWeekly()
        {
            return Ok(MSWeeklyMessageStat.GetWeeklyMessageStats("AZURESERVICEHEALTHALERT"));
        }

        [Route("api/ado/projects")]
        [HttpGet]
        public IActionResult GetADOProjects()
        {
            string personalAccessToken = GlobalConfiguration.Instance.GetSecret("AzureDevOpsPAT");
            VssConnection connection = new VssConnection(new Uri("https://dev.azure.com/adritberatung"), new VssBasicCredential(string.Empty, personalAccessToken));
            ProjectHttpClient client = connection.GetClient<ProjectHttpClient>();

            return Ok(client.GetProjects().Result);
        }

        [Route("api/ado/projects/{id}/properties")]
        [HttpGet]
        public IActionResult GetADOProjectProperties(Guid id)
        {
            string personalAccessToken = GlobalConfiguration.Instance.GetSecret("AzureDevOpsPAT");
            VssConnection connection = new VssConnection(new Uri("https://dev.azure.com/adritberatung"), new VssBasicCredential(string.Empty, personalAccessToken));
            ProjectHttpClient client = connection.GetClient<ProjectHttpClient>();

            return Ok(client.GetProjectPropertiesAsync(id).Result);
        }

        [Route("api/ado/process/{id}")]
        [HttpGet]
        public IActionResult GetADOProcessDefinition(Guid id)
        {
            string personalAccessToken = GlobalConfiguration.Instance.GetSecret("AzureDevOpsPAT");
            VssConnection connection = new VssConnection(new Uri("https://dev.azure.com/adritberatung"), new VssBasicCredential(string.Empty, personalAccessToken));
            WorkItemTrackingProcessHttpClient client = connection.GetClient<WorkItemTrackingProcessHttpClient>();

            return Ok(client.GetProcessWorkItemTypesAsync(id).Result);
        }

        [Route("api/ado/process/{id}/workItemType/{witRefName}")]
        [HttpGet]
        public IActionResult GetADOWorkItemType(Guid id, string witRefName)
        {
            string personalAccessToken = GlobalConfiguration.Instance.GetSecret("AzureDevOpsPAT");
            VssConnection connection = new VssConnection(new Uri("https://dev.azure.com/adritberatung"), new VssBasicCredential(string.Empty, personalAccessToken));
            WorkItemTrackingProcessHttpClient client = connection.GetClient<WorkItemTrackingProcessHttpClient>();

            var result = client.GetAllWorkItemTypeFieldsAsync(id, witRefName).Result;
            return Ok(result);
        }
    }

    [ApiController]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    public class AzureUpdate : ControllerBase
    {
        private readonly ILogger<AzureUpdate> _logger;

        public AzureUpdate(ILogger<AzureUpdate> logger)
        {
            _logger = logger;
        }

        // Azure Health Alerts endpoints
        [Route("api/azure/updates")]
        [HttpGet]
        public IActionResult GetAzureUpdates(bool? active)
        {
            var db = new MSSHNotificationDatabase();

            MSUserInfo userInfo = new(User);
            if (!Guid.TryParse(userInfo.ObjectId, out Guid userId)) { userId = Guid.Empty; }

            return Ok(db.GetAzureUpdates(active, userId));
        }

        [Route("api/azure/updates/{id}")]
        [HttpGet]
        public IActionResult GetAzureUpdate(string id)
        {
            var db = new MSSHNotificationDatabase();

            MSUserInfo userInfo = new(User);
            if (!Guid.TryParse(userInfo.ObjectId, out Guid userId)) { userId = Guid.Empty; }

            if (userId != Guid.Empty)
                db.SetViewpointReadFlag(userId, id, Guid.Empty, Guid.Empty, true);

            return Ok(db.GetAzureUpdate(id, userId));
        }
    }
}
