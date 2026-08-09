using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Authorization;
using Microsoft.ServiceHealthHub.Core;

namespace Service_Health_Dashboard_v2.Controllers
{
    [Authorize(Roles = "Admin")]
    [ApiController]
    public class Settings : ControllerBase
    {
        private readonly ILogger<PublicReport> _logger;

        public Settings(ILogger<PublicReport> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("api/[controller]/application")]
        public IActionResult Get()
        {
            MSApplicationSettings config = new()
            {
                pirStorageUri = GlobalConfiguration.Instance.GetConfigurationValue(Constants.c_settings_PostIncidentReview_Storage),
                releaseMessageEndpoint = GlobalConfiguration.Instance.GetConfigurationValue(Constants.c_settings_ServiceHealthHub_ReleaseNotesEndpoint),
                azureTranslatorEnabled = GlobalConfiguration.Instance.AzureTranslatorEnabled,
                azureTranslatorLocation = GlobalConfiguration.Instance.AzureTranslatorResourceLocation,
                azureTranslatorKey = GlobalConfiguration.Instance.AzureTranslatorSubscriptionKey,
                languageServiceEnabled = GlobalConfiguration.Instance.LanguageServiceEnabled,
                languageServiceEndpoint = GlobalConfiguration.Instance.LanguageServiceEndpoint,
                languageServiceKey = GlobalConfiguration.Instance.LanguageServiceSubscriptionKey,
                textSummarizationSentenceCount = GlobalConfiguration.Instance.TextSummarizationSentenceCount,
                branding = GlobalConfiguration.Instance.Branding
            };

            return Ok(config);
        }

        [HttpPost]
        [Route("api/[controller]/application")]
        public IActionResult Post([FromBody] MSApplicationSettings body)
        {
            GlobalConfiguration.Instance.SetConfigurationValue(Constants.c_settings_PostIncidentReview_Storage, body.pirStorageUri);
            GlobalConfiguration.Instance.SetConfigurationValue(Constants.c_settings_ServiceHealthHub_ReleaseNotesEndpoint, body.releaseMessageEndpoint);
            GlobalConfiguration.Instance.AzureTranslatorEnabled = body.azureTranslatorEnabled;
            GlobalConfiguration.Instance.AzureTranslatorResourceLocation = body.azureTranslatorLocation;
            GlobalConfiguration.Instance.AzureTranslatorSubscriptionKey = body.azureTranslatorKey;
            GlobalConfiguration.Instance.LanguageServiceEnabled = body.languageServiceEnabled;
            GlobalConfiguration.Instance.LanguageServiceEndpoint = body.languageServiceEndpoint;
            GlobalConfiguration.Instance.LanguageServiceSubscriptionKey = body.languageServiceKey;
            GlobalConfiguration.Instance.TextSummarizationSentenceCount = body.textSummarizationSentenceCount;
            GlobalConfiguration.Instance.Branding = body.branding != null ? body.branding : new MSApplicationBrandingSettings();
            return Ok();
        }
    }
}
