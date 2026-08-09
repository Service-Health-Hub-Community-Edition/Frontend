using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core.Graph;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{

    [Authorize(Roles = "Admin")]
    [ApiController]
    public class Search : ControllerBase
    {
        private readonly ILogger<Search> _logger;

        public Search(ILogger<Search> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("api/admin/[controller]")]
        public IActionResult Get()
        {

            MSSearchSettings settings = new();

            return Ok(settings);
        }

        [HttpPost]
        [Route("api/admin/[controller]")]
        public async Task<IActionResult> Post([FromBody] SearchConfigRequest body)
        {
            MSSearchSettings settings = new();

            if (body.rootUrl != null && body.rootUrl.Trim() == "")
            {
                body.rootUrl = HttpContext.Request.Scheme + "://" + HttpContext.Request.Host;
                // TODO: set reindex flag
            }

            switch (body.operation)
            {
                case SearchOperation.create:
                    settings.Configure();
                    settings.ConnectorSettings.SetConfig(body);
                    break;
                case SearchOperation.configure:                  
                    settings.ConnectorSettings.SetConfig(body);
                    break;
                case SearchOperation.enable:
                    settings.ConnectorSettings.Enable();
                    break;
                case SearchOperation.disable:
                    settings.ConnectorSettings.Disable();
                    break;
                case SearchOperation.schemaUpdate:
                    await settings.ConfigureSchema(settings.GraphAPIConnector);
                    break;
                default:
                    break;
            }
            
            return Ok(settings);
        }
    }
}
