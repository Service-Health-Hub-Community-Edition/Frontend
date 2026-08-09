using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Microsoft.ServiceHealthHub.Controllers
{
    public class AuditLogEntry
    {
        public Guid Id { get; set; }
        public DateTime Timestamp { get; set; }
        public Guid CorrelationId { get; set; }
        public string User { get; set; }
        public string Activity { get; set; }
        public string Item { get; set; }
        public string ItemType { get; set; }
        public string ItemUniqueId { get; set; }
        public string EventSource { get; set; }
        public dynamic ExtendedProperties { get; set; }
        public dynamic ModifiedProperties { get; set; }  
    }

    public class AuditLogCollection
    {
        private static AuditLogEntry InitializeAuditLogEntry(MSPropertyBag instanceData)
        {
            return new AuditLogEntry
            {
                Id = instanceData["ID"] is DBNull ? Guid.Empty : (Guid)instanceData["ID"],
                Timestamp = instanceData["Timestamp"] is DBNull ? DateTime.MinValue : (DateTime)instanceData["Timestamp"],
                CorrelationId = instanceData["CorrelationID"] is DBNull ? Guid.Empty : (Guid)instanceData["CorrelationID"],
                User = instanceData["User"] is DBNull ? null : (string)instanceData["User"],
                Activity = instanceData["Activity"] is DBNull ? string.Empty : (string)instanceData["Activity"],
                Item = instanceData["Item"] is DBNull ? string.Empty : (string)instanceData["Item"],
                ItemType = instanceData["ItemType"] is DBNull ? string.Empty : (string)instanceData["ItemType"],
                ItemUniqueId = instanceData["ItemUniqueId"] is DBNull ? string.Empty : (string)instanceData["ItemUniqueId"],
                EventSource = instanceData["EventSource"] is DBNull ? string.Empty : (string)instanceData["EventSource"],
                ExtendedProperties = instanceData["ExtendedProperties"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ExtendedProperties"]),
                ModifiedProperties = instanceData["ModifiedProperties"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ModifiedProperties"])
            };
        }
        public static List<AuditLogEntry> GetAuditLogEntries(string scheme, string id, string type)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            List<AuditLogEntry> auditLogEntries = new List<AuditLogEntry>();

            List<MSPropertyBag> dbResult = db.GetAuditLog(string.Format("{0}://{1}/{2}", scheme, type, id));
            foreach (MSPropertyBag instanceData in dbResult)
            {
                auditLogEntries.Add(InitializeAuditLogEntry(instanceData));
            }

            return auditLogEntries;
        }

        public static List<AuditLogEntry> GetAuditLogEntries(Guid correlationId)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            List<AuditLogEntry> auditLogEntries = new List<AuditLogEntry>();

            List<MSPropertyBag> dbResult = db.GetAuditLog(correlationId);
            foreach (MSPropertyBag instanceData in dbResult)
            {
                auditLogEntries.Add(InitializeAuditLogEntry(instanceData));
            }

            return auditLogEntries;
        }
    }

    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    [ApiController]
    public class AuditLog : ControllerBase
    {
        private readonly ILogger<CustomAction> _logger;

        public AuditLog(ILogger<CustomAction> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<List<AuditLogEntry>> GetAsync(Guid? correlationId, string scheme, string id, string type)
        {
            if (correlationId != null)
                return AuditLogCollection.GetAuditLogEntries(correlationId.Value);
            else
                return AuditLogCollection.GetAuditLogEntries(scheme, id, type);
        }
    }
}
