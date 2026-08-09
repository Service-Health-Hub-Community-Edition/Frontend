using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSServiceHealthDashboardConfig
    {
        internal string m_ClientId = string.Empty;
        internal string m_TenantId = string.Empty;
        internal string m_DBVersion = string.Empty;
        internal string m_AppTitle = string.Empty;
        internal MSApplicationBrandingSettings m_branding = new MSApplicationBrandingSettings();
        
        public string clientId => m_ClientId;
        public string tenantId => m_TenantId;
        public string dbVersion => m_DBVersion;
        public string appTitle => m_AppTitle;
        public MSApplicationBrandingSettings branding => m_branding;

        public MSServiceHealthDashboardConfig()
        {
            m_ClientId = GlobalConfiguration.Instance.ClientAppId;
            m_TenantId = string.IsNullOrWhiteSpace(GlobalConfiguration.Instance.ClientTenantDomain) ? 
                GlobalConfiguration.Instance.TenantDomain :
                GlobalConfiguration.Instance.ClientTenantDomain;

            MSSHNotificationDatabase db = new();
            m_DBVersion = db.GetSchemaVersion();
            m_branding = GlobalConfiguration.Instance.Branding;
            m_AppTitle = GlobalConfiguration.Instance.ApplicationTitle;
        }
    }

    [Route("api/[controller]")]
    [ApiController]
    public class Config : ControllerBase
    {
        private readonly ILogger<Config> _logger;

        public Config(ILogger<Config> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task<MSServiceHealthDashboardConfig> Get()
        {
            MSServiceHealthDashboardConfig config = new MSServiceHealthDashboardConfig();
            return config;
        }
    }
}
