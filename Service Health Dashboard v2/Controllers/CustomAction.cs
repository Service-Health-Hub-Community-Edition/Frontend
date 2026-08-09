using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.ServiceHealthHub.Core;

using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;

namespace Microsoft.ServiceHealthHub.Controllers
{
    public class CustomActionData
    {
        public Guid actionId = Guid.Empty;
        public string communicationType = string.Empty;
        public List<string> communicationIds = new List<string>();
    }

    public class CustomActionAuditResult
    {
        public CustomActionResponse actionResponse;
        public string actionName;
    }

    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Admin")]
    [ApiController]
    public class CustomAction : ControllerBase
    {
        private readonly ILogger<CustomAction> _logger;

        public CustomAction(ILogger<CustomAction> logger)
        {
            _logger = logger;
        }

        [Route("api/[controller]")]
        [HttpGet]
        public async Task<List<CustomBaseAction>> GetAsync(Guid? componentId, string componentType)
        {
            Guid cid = componentId.GetValueOrDefault(Guid.Empty);

            if (cid != Guid.Empty)
                return CustomBaseAction.GetComponentInstances(cid);
            else if (!string.IsNullOrWhiteSpace(componentType))
                return CustomBaseAction.GetComponentInstances(componentType);
            else
                return new List<CustomBaseAction>();
        }

        [Authorize(Roles = "Admin")]
        [Route("api/[controller]")]
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] CustomBaseActionInt body)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.UpdateCustomAction(body);
            return Ok();
        }

        [Authorize(Roles = "Admin")]
        [Route("api/[controller]")]
        [HttpDelete]
        public async Task<IActionResult> DeleteAsync([FromBody] CustomBaseActionInt body)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.DeleteCustomAction(body);
            return Ok();
        }

        [Route("api/[controller]/[action]")]
        [HttpGet]
        public async Task<List<CustomActionType>> ActionTypeAsync()
        {
            return CustomActionType.GetCustomActionTypes();
        }

        [Route("api/[controller]/[action]")]
        [HttpPost]
        public async Task<IActionResult> RunAsync([FromBody] CustomActionData body)
        {
            Dictionary<string, CustomActionResponse> response = new Dictionary<string, CustomActionResponse>();
            MSSHNotificationDatabase database = new MSSHNotificationDatabase();
            MSUserInfo userInfo = new MSUserInfo(User);

            CustomBaseAction action = CustomBaseAction.GetInstance(body.actionId);

            foreach (string communicationId in body.communicationIds)
            {
                if (action == null)
                {
                    CustomActionResponse res = new CustomActionResponse
                    {
                        IsSuccessStatusCode = false,
                        ResponseBody = string.Format("Action with id '{0}' not found. Please contact your Service Health Hub administrator.", body.actionId),
                        StatusCode = System.Net.HttpStatusCode.NotFound
                    };

                    response.Add(communicationId, res);

                    database.AddAuditLogRecord(Guid.Empty, Guid.Empty, "", userInfo.UserName, "Failed",
                        string.Format("item://{0}/{1}", body.communicationType, communicationId), "Action",
                        Guid.Empty, communicationId, "WebApp", Guid.Empty,
                            new CustomActionAuditResult()
                            {
                                actionName = "Not found",
                                actionResponse = res
                            }, null);

                }
                else
                {
                    CommunicationBase comm = null;
                    CustomActionMessage msg;
                    CustomActionResponse res;

                    bool unsupported = true;

                    switch (action.Component.InternalName.ToString().ToLower())
                    {
                        case "serviceupdatemessage":
                            comm = new M365MessageCenterCommunication(communicationId);
                            unsupported = false;
                            break;

                        case "servicehealthissue":
                            comm = new M365ServiceHealthIssue(communicationId);
                            unsupported = false;
                            break;

                        case "roadmapcommunication":
                            comm = new M365Roadmap(communicationId);
                            unsupported = false;
                            break;

                        case "azureservicehealthalert":
                            comm = new CommunicationBase();
                            MSPropertyBag dbComm = database.GetAzureHealthAlert(communicationId);
                            comm.WorkItemId = dbComm.ContainsKey("WorkItemID") ? dbComm["WorkItemID"].ToString() : null;
                            comm.WorkItemUrl = dbComm.ContainsKey("WorkItemURL") ? dbComm["WorkItemURL"].ToString() : null;
                            if (dbComm.ContainsKey("Data"))
                            {
                                object data = JsonSerializer.Deserialize<object>((string)dbComm["Data"]);
                                dbComm["Data"] = data;
                            }
                            comm.Message = dbComm;
                            unsupported = false;
                            break;

                        case "azureupdate":
                            comm = new CommunicationBase();
                            MSPropertyBag dbCommAZU = database.GetAzureUpdate(communicationId);
                            comm.WorkItemId = dbCommAZU.ContainsKey("WorkItemID") ? dbCommAZU["WorkItemID"].ToString() : null;
                            comm.WorkItemUrl = dbCommAZU.ContainsKey("WorkItemURL") ? dbCommAZU["WorkItemURL"].ToString() : null;
                            if (dbCommAZU.ContainsKey("Data"))
                            {
                                object data = JsonSerializer.Deserialize<object>((string)dbCommAZU["Data"]);
                                dbCommAZU["Data"] = data;
                            }
                            comm.Message = dbCommAZU;
                            unsupported = false;
                            break;

                        case "d365powerplatformrelease":
                            comm = new CommunicationBase();
                            MSPropertyBag dbCommd365pp = database.GetD365PowerPlatformRelease(communicationId);
                            comm.WorkItemId = dbCommd365pp.ContainsKey("WorkItemID") ? dbCommd365pp["WorkItemID"].ToString() : null;
                            comm.WorkItemUrl = dbCommd365pp.ContainsKey("WorkItemURL") ? dbCommd365pp["WorkItemURL"].ToString() : null;
                            if (dbCommd365pp.ContainsKey("Data"))
                            {
                                object data = JsonSerializer.Deserialize<object>((string)dbCommd365pp["Data"]);
                                dbCommd365pp["Data"] = data;
                            }
                            comm.Message = dbCommd365pp;
                            unsupported = false;
                            break;

                        default:
                            break;
                    }

                    if (!unsupported)
                    {
                        msg = new CustomActionMessage
                        {
                            WorkItemId = string.IsNullOrWhiteSpace(comm.WorkItemId) ? "" : comm.WorkItemId,
                            WorkItemUrl = string.IsNullOrWhiteSpace(comm.WorkItemUrl) ? "" : comm.WorkItemUrl,
                            Message = comm.Message
                        };
                        res = action.Run(msg);

                        response.Add(communicationId, res);

                        CustomActionResponse auditLogRes = new CustomActionResponse() {
                            IsSuccessStatusCode = res.IsSuccessStatusCode,
                            Response = null,
                            ResponseBody = res.ResponseBody,
                            StatusCode = res.StatusCode
                        };
                        
                        if (res.IsSuccessStatusCode)
                            database.AddAuditLogRecord(Guid.Empty, Guid.Empty, "", userInfo.UserName, "Success",
                                string.Format("item://{0}/{1}", body.communicationType, communicationId), "Action",
                                Guid.Empty, communicationId, "WebApp", Guid.Empty,
                                new CustomActionAuditResult()
                                {
                                    actionName = action.Name,
                                    actionResponse = auditLogRes
                                }, null);
                        else
                            database.AddAuditLogRecord(Guid.Empty, Guid.Empty, "", userInfo.UserName, "Failed",
                                string.Format("item://{0}/{1}", body.communicationType, communicationId), "Action",
                                Guid.Empty, communicationId, "WebApp", Guid.Empty,
                                new CustomActionAuditResult()
                                {
                                    actionName = action.Name,
                                    actionResponse = auditLogRes
                                }, null);
                    }
                    else
                    {
                        CustomActionResponse unsupportedRes = new CustomActionResponse
                        {
                            IsSuccessStatusCode = false,
                            ResponseBody = string.Format("Cannot run action. Unsupported communication type for communication id {0}.", communicationId),
                            StatusCode = System.Net.HttpStatusCode.NotImplemented
                        };

                        response.Add(communicationId, unsupportedRes);

                        database.AddAuditLogRecord(Guid.Empty, Guid.Empty, "", userInfo.UserName, "Failed",
                            string.Format("item://{0}/{1}", body.communicationType, communicationId), "Action",
                            Guid.Empty, communicationId, "WebApp", Guid.Empty,
                            new CustomActionAuditResult()
                            {
                                actionName = action.Name,
                                actionResponse = unsupportedRes
                            }, null);
                    }
                }
            }

            return Ok(response);
        }
    }
}
