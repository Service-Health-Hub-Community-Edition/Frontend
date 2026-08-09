using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSUserProfileSetPropertyResponse
    {
        public int result { get; set; }
    }

    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,LicenseReader,Public,Admin")]
    [ApiController]
    public class Users : ControllerBase
    {
        private readonly ILogger<Items> _logger;

        public Users(ILogger<Items> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("api/[controller]/{id?}")]
        public async Task<IActionResult> Get(Guid? id)
        {
            if (id != null)
            {
                dynamic res = new ExpandoObject();
                MSSHNotificationDatabase db = new MSSHNotificationDatabase();
                List<MSPropertyBag> dbResult = db.GetUserProfile(id.Value);
                
                if (dbResult.Count <= 0)
                    return NotFound(string.Format("User with object id {0} is not found.", id.Value));
                
                MSPropertyBag user = dbResult[0];

                res.id = user["ObjectId"] is DBNull ? Guid.Empty : (Guid)user["ObjectId"];
                res.properties = user["Properties"] is DBNull ? null : JsonConvert.DeserializeObject((string)user["Properties"]);
                if (!(user["LastModified"] is DBNull))
                    res.lastModified = (DateTime)user["LastModified"];

                return Ok(res);
            } else {
                return BadRequest("User object id is not provided.");
            }
        }

        [HttpPost]
        [Route("api/[controller]/{id?}")]
        public async Task<IActionResult> PostAsync(Guid? id, [FromBody] dynamic body)
        {
            if (id != null)
            {
                MSSHNotificationDatabase db = new MSSHNotificationDatabase();
                string json = JsonConvert.SerializeObject(body);

                MSUserProfileSetPropertyResponse res = new MSUserProfileSetPropertyResponse()
                {
                    result = db.SetUserProfileProperties(id.Value, json)
                };

                return Ok(res);
            } else
            {
                return BadRequest("User object id is not provided.");
            }
        }
    }
}
