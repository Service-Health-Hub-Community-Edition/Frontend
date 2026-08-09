using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Azure;
using Microsoft.ServiceHealthHub.Core;
using System.Collections.Generic;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSTranslationRequest
    {
        public string language = string.Empty;
        public string[] contents;
    }

    [Route("api/[controller]")]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Public,Admin,LicenseReader")]
    [ApiController]
    public class Translate : ControllerBase
    {
        private readonly ILogger<PublicReport> _logger;

        public Translate(ILogger<PublicReport> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        public Dictionary<string, string> Post([FromBody] MSTranslationRequest body)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            if (string.IsNullOrWhiteSpace(body.language) || body.language.Trim().ToLower() == "en")
            {
                // copy original values and return
                foreach (string content in body.contents)
                {
                    result.Add(content, content);
                }
            }
            else
            {
                // translate and return
                List<MSTranslationCollection> translationResult = MSAzureTranslator.Translate(body.contents, body.language);

                foreach (string content in body.contents)
                {
                    MSTranslationCollection contentTranslation = translationResult.Find(translation => translation.OriginalMessage == content);
                    if (string.IsNullOrWhiteSpace(contentTranslation?.Translations?[body.language]?.Message))
                    {
                        result.Add(content, content);
                    } else
                    {
                        result.Add(content, contentTranslation?.Translations?[body.language]?.Message);
                    }
                }
            }

            return result;
        }
    }
}
