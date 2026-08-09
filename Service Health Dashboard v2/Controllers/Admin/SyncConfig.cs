using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.AspNetCore.Authorization;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSSyncConfig
    {
        public string component = string.Empty;
        public string element = string.Empty;
        public string config = string.Empty;
    }

    [Route("api/[controller]")]
    [Authorize(Roles = "Admin")]
    [ApiController]
    public class SyncConfig : ControllerBase
    {
        private readonly ILogger<PublicReport> _logger;

        public SyncConfig(ILogger<PublicReport> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task<MSSyncConfigEntry> Get(string component, string element)
        {
            MSSyncConfigEntry configEntry = new MSSyncConfigEntry(component.Trim().ToUpper(), element.Trim().ToUpper());
            return configEntry;
        }

        [HttpPost]
        public async Task<IActionResult> Post([FromBody] MSSyncConfig body)
        {
            MSUserInfo userInfo = new MSUserInfo(User);
            MSSyncConfigEntry configEntry = new MSSyncConfigEntry(body.component, body.config);
            configEntry.SetConfigEntry(body.component, body.element, body.config);
            configEntry.Update(userInfo);
            return Ok();
        }
    }
}
