using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Azure;
using Microsoft.ServiceHealthHub.Core;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSPublicMessagePublishData
    {
        public string id = string.Empty;
        public string type = string.Empty;
        public string comments = string.Empty;
    }

    public class MSServicePublicEvents
    {
        internal MSPublicIncidentCollection _events = new MSPublicIncidentCollection();
        internal MSPublicMessageCollection _messages = new MSPublicMessageCollection();
        public MSPublicIncidentCollection Events => _events;
        public MSPublicMessageCollection Messages => _messages;

        public MSServicePublicEvents()
        {
            _events = MSPublicIncident.GetPublicIncidents();
            _messages = MSPublicMessage.GetPublicMessages();
        }

        public MSServicePublicEvents(string lang)
        {
            MSPublicIncidentCollection events = MSPublicIncident.GetPublicIncidents();

            foreach (MSPublicIncident incident in events)
            {
                string[] messages = new string[] { incident.Title, incident.Comments };

                List<MSTranslationCollection> result = MSAzureTranslator.Translate(messages, lang);

                MSTranslationCollection titleTranslation = result.Find(translation => translation.OriginalMessage == incident.Title);
                MSTranslationCollection commentsTranslation = result.Find(translation => translation.OriginalMessage == incident.Comments);

                if (!string.IsNullOrWhiteSpace(titleTranslation?.Translations?[lang]?.Message) && !string.IsNullOrWhiteSpace(commentsTranslation?.Translations?[lang]?.Message))
                {
                    incident.SetPublicIncident(titleTranslation?.Translations?[lang]?.Message, commentsTranslation?.Translations?[lang]?.Message);
                }
            }

            MSPublicMessageCollection messageCollection = MSPublicMessage.GetPublicMessages();

            foreach (MSPublicMessage message in messageCollection)
            {
                string[] messages = new string[] { message.Title, message.Comments };

                List<MSTranslationCollection> result = MSAzureTranslator.Translate(messages, lang, true);

                MSTranslationCollection titleTranslation = result.Find(translation => translation.OriginalMessage == message.Title);
                MSTranslationCollection commentsTranslation = result.Find(translation => translation.OriginalMessage == message.Comments);
                // MSTranslationCollection contentTranslation = result.Find(translation => translation.OriginalMessage == message.Content);

                if (!string.IsNullOrWhiteSpace(titleTranslation?.Translations?[lang]?.Message))
                {
                    message.SetPublicMessage(
                        titleTranslation?.Translations?[lang]?.Message, 
                        commentsTranslation?.Translations?[lang]?.Message);
                }
            }

            _events = events;
            _messages = messageCollection;
        }
    }

    [Route("api/[controller]")]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Public,Admin,LicenseReader")]
    [ApiController]
    public class PublicReport : ControllerBase
    {
        private readonly ILogger<PublicReport> _logger;

        public PublicReport(ILogger<PublicReport> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task<MSServicePublicEvents> Get(string lang)
        {
            if (string.IsNullOrWhiteSpace(lang))
            {
                return new MSServicePublicEvents();
            }
            else
            {
                return new MSServicePublicEvents(lang);
            }
        }
    }

    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Public,Admin,LicenseReader")]
    [ApiController]
    public class PublicMessage : ControllerBase
    {
        private readonly ILogger<PublicReport> _logger;

        public PublicMessage(ILogger<PublicReport> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<MSPublicMessage> Get(string id, string lang)
        {
            return MSPublicMessage.GetPublicMessage(id, lang);
        }

        [Route("api/[controller]/[action]")]
        [HttpPost]
        public async Task<IActionResult> PublishAsync([FromBody] MSPublicMessagePublishData body)
        {
            MSUserInfo userInfo = new MSUserInfo(User);
            MSPublicMessage.Publish(body.id, body.type, body.comments, userInfo);
            return Ok();
        }

        [Route("api/[controller]/[action]")]
        [HttpPost]
        public async Task<IActionResult> UnpublishAsync([FromBody] MSPublicMessagePublishData body)
        {
            MSUserInfo userInfo = new MSUserInfo(User);
            MSPublicMessage.Unpublish(body.id, body.type, userInfo);
            return Ok();
        }
    }
}
