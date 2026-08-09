using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class O365EndpointSet
    {
        public int id  { get; set; }
        public string serviceArea { get; set; }
        public string serviceAreaDisplayName { get; set; }
        public string category { get; set; }
        public bool expressRoute { get; set; }
        public bool required { get; set; }
        public string[] ips { get; set; }
        public string[] urls { get; set; }
        public string tcpPorts { get; set; }
        public string udpPorts { get; set; }
        public string notes { get; set; }
        public List<O365ChangeRecord> changes { get; set; }
    }

    public class O365EndpointAdditionChange
    {
        public string effectiveDate { get; set; }
        public string[] ips { get; set; }
        public string[] urls { get; set; }
    }

    public class O365EndpointRemovalChange
    {
        public string[] ips { get; set; }
        public string[] urls { get; set; }
    }

    public class O365ChangeRecord
    {
        public int id { get; set; }
        public int endpointSetId { get; set; }
        public string disposition { get; set; }
        public string impact { get; set; }
        public string version { get; set; }
        public dynamic previous { get; set; }
        public dynamic current { get; set; }
        public dynamic add { get; set; }
        public dynamic remove { get; set; }
    }

    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    [ApiController]
    public class EndpointSets : ControllerBase
    {
        private readonly ILogger<Items> _logger;

        private string Get(string uri)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public EndpointSets(ILogger<Items> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("api/[controller]")]
        public async Task<IActionResult> Get()
        {
            Guid requestId = Guid.NewGuid();
            string endpointSetsStr = Get(string.Format("https://endpoints.office.com/endpoints/worldwide?clientrequestid={0}", requestId));
            string changesStr = Get(string.Format("https://endpoints.office.com/changes/worldwide/0000000000?clientrequestid={0}", requestId));

            List<O365EndpointSet> endpointSets = JsonConvert.DeserializeObject<List<O365EndpointSet>>(endpointSetsStr);
            List<O365ChangeRecord> changeRecords = JsonConvert.DeserializeObject<List<O365ChangeRecord>>(changesStr);

            foreach (O365EndpointSet endpointSet in endpointSets)
                endpointSet.changes = changeRecords.FindAll((O365ChangeRecord changeRec) => changeRec.endpointSetId == endpointSet.id);

            return Ok(endpointSets);
        }

        /* [HttpGet]
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
        } */
    }
}
