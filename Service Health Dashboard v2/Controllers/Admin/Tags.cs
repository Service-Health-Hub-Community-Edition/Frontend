using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.TeamFoundation;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSGroupedTagDefinition
    {
        public string Component { get; set; }
        public string InternalName { get; set; }
        public List<MSTagDefinition> TagDefinitions { get; set; }
    }

    public class MSTagDefinitionOperationData
    {
        public string? name { get; set; }
        public string? targetGroup { get; set; }
    }

    public class MSTagDefinitionOperation
    {
        public string op { get; set; }
        public MSTagDefinitionOperationData data { get; set; }

    }

    public class MSTagOperation
    {
        public string op { get; set; }
        public Guid tagId { get; set; }
    }

    // [Authorize(Roles = "Admin")]
    [ApiController]
    public class TagDefinitions : ControllerBase
    {
        private readonly ILogger<TagDefinitions> _logger;

        public TagDefinitions(ILogger<TagDefinitions> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<List<MSGroupedTagDefinition>> Get(string? componentName)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            List<MSGroupedTagDefinition> res = new();

            MSTagDefinitions def = MSTagDefinitions.CreateInstance();
            List<MSComponent> components = db.GetComponent((Guid?)null);
            foreach (MSComponent component in components)
            {
                if (component.Capabilities.Contains("Portal"))
                {
                    if (!(!string.IsNullOrWhiteSpace(componentName) && component.InternalName.ToUpper() != componentName.ToUpper())) 
                    {
                        MSGroupedTagDefinition groupedTagDefinition = new()
                        {
                            Component = component.Name,
                            InternalName = component.InternalName,
                            TagDefinitions = def.FindAll(d => d.Type?.ToUpper() == component.InternalName.ToUpper())
                        };

                        res.Add(groupedTagDefinition);
                    }
                }
            }

            MSGroupedTagDefinition general = new()
            {
                Component = "General",
                InternalName = null,
                TagDefinitions = def.FindAll(d => d.Type == null)
            };

            res.Add(general);

            return res;
        }

        [Route("api/[controller]/{id}")]
        [Authorize(Roles = "Admin")]
        [HttpPost]
        public async Task<StatusCodeResult> Post(Guid? id, [FromBody] MSTagDefinitionOperation data)
        {
            if (data == null)
                return new BadRequestResult();

            MSTagDefinition def;

            switch (data.op.ToLower())
            {
                case "move":
                    if (id != null)
                    {
                        def = new(id.Value);
                        def.Move(data.data.targetGroup);
                    } 
                    else
                    {
                        return new NotFoundResult();
                    }
                    break;
                case "rename":
                    if (id != null)
                    {
                        def = new(id.Value);
                        def.Name = data.data.name;
                        def.Update();
                    }
                    else
                    {
                        return new NotFoundResult();
                    }
                    break;
                default:
                    return new BadRequestResult();
            }
            
            return new OkResult();
        }

        [Route("api/[controller]")]
        [Authorize(Roles = "Admin")]
        [HttpPost]
        public async Task<StatusCodeResult> Post([FromBody] MSTagDefinitionOperation data)
        {
            if (data == null)
                return new BadRequestResult();

            MSTagDefinition def;

            switch (data.op.ToLower())
            {
                case "create":
                    MSTagDefinition newDef = new(data.data.name, data.data.targetGroup);
                    newDef.Update();
                    break;
                default:
                    return new BadRequestResult();
            }

            return new OkResult();
        }

        [Route("api/[controller]/{id}")]
        [Authorize(Roles = "Admin")]
        [HttpDelete]
        public async Task<StatusCodeResult> Delete(Guid id)
        {
            MSTagDefinition def = new(id);

            def.Delete();

            return new OkResult();
        }
    }

    [ApiController]
    public class Tag : ControllerBase
    {
        private readonly ILogger<TagDefinitions> _logger;

        public Tag(ILogger<TagDefinitions> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]/{type}/{messageId}")]
        [HttpGet]
        public async Task<MSTags> Get(string messageId, string type)
        {
            MSTags res = MSTags.CreateInstance(messageId, type);

            return res;
        }

        [Route("api/[controller]/{type}/{messageId}")]
        [Authorize(Roles = "Admin,Communication.Write.All")]
        [HttpPost]
        public async Task<StatusCodeResult> Post(string messageId, string type, [FromBody] MSTagOperation data)
        {
            if (data == null || string.IsNullOrWhiteSpace(messageId) || string.IsNullOrWhiteSpace(type))
                return new BadRequestResult();

            switch (data.op.ToLower())
            {
                case "add":
                    MSServiceNotification communication = new();
                    MSSHNotificationDatabase db = new();
                    db.GetNotification(messageId, type, communication);

                    if (string.IsNullOrWhiteSpace(communication.Id))
                        return new NotFoundResult();

                    MSTag tag = new(messageId, type, data.tagId);
                    tag.Update();

                    break;
                default:
                    return new BadRequestResult();
            }

            return new OkResult();
        }

        [Route("api/[controller]/{type}/{messageId}/{tagId}")]
        [Authorize(Roles = "Admin,Communication.Write.All")]
        [HttpDelete]
        public async Task<StatusCodeResult> Delete(string messageId, string type, Guid tagId)
        {
            MSTag? tag = MSTag.CreateInstance(messageId, type, tagId);
            if (tag == null)
                return new NotFoundResult();

            tag.Delete();

            return new OkResult();
        }
    }
}
