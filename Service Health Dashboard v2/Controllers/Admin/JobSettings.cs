using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Authorization;
using Microsoft.ServiceHealthHub.Core;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSJobSettings
    {
        public bool enabled;
    }

    [Authorize(Roles = "Admin")]
    [ApiController]
    public class JobSettings : ControllerBase
    {
        private readonly ILogger<JobSettings> _logger;

        public JobSettings(ILogger<JobSettings> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("api/jobs/{id}/settings")]
        public IActionResult Get(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return NotFound();

            MSSHNotificationDatabase db = new();

            MSJobSettings config = db.GetConfigurationValue<MSJobSettings>("Config-" + id);
            if (config == null)
                config = new()
                {
                    enabled = false
                };

            return Ok(config);
        }

        [HttpPost]
        [Route("api/jobs/{id}/settings")]
        public IActionResult Post(string id, [FromBody] MSJobSettings body)
        {
            if (string.IsNullOrWhiteSpace(id))
                return NotFound();

            MSSHNotificationDatabase db = new();
            db.SetConfigurationValue("Config-" + id, body);
            return Ok();
        }
    }
}
