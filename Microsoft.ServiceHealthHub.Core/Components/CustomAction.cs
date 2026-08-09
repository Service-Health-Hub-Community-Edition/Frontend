using System.Net;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using Azure.Messaging.ServiceBus;

namespace Microsoft.ServiceHealthHub.Core
{
    public class CustomActionTypeOptionParameter
    {
        public string internalName { get; set; }
        public string displayName { get; set; }
        public bool secureString { get; set; }

    }

    public class CustomActionTypeOption
    {
        public string internalName { get; set; }
        public string displayName { get; set; }
        public List<CustomActionTypeOptionParameter> parameters { get; set; }
    }

    public class CustomActionTypeParameter
    {
        public string internalName { get; set; }
        public string displayName { get; set; }
        public string type { get; set; }
        public List<CustomActionTypeOption> options { get; set; }

    }

    public class CustomActionType
    {
        MSSHNotificationDatabase _db = new MSSHNotificationDatabase();

        Guid id;
        string name;
        string icon;
        List<CustomActionTypeParameter> parameters;

        public Guid Id { get { return id; } }
        public string Name { get { return name; } }
        public string Icon { get { return icon; } }
        public List<CustomActionTypeParameter> Parameters { get { return parameters; } }

        public static List<CustomActionType>GetCustomActionTypes()
        {
            MSSHNotificationDatabase _db = new MSSHNotificationDatabase();
            List<MSPropertyBag> dbResultSet = _db.GetCustomActionTypes();
            List<CustomActionType> result = new List<CustomActionType>();

            foreach (MSPropertyBag dbResult in dbResultSet)
                result.Add(new CustomActionType(
                        dbResult["ActionTypeId"] is DBNull ? Guid.Empty : (Guid)dbResult["ActionTypeId"],
                        dbResult["Name"] is DBNull ? string.Empty : (string)dbResult["Name"],
                        dbResult["Icon"] is DBNull ? string.Empty : (string)dbResult["Icon"],
                        dbResult["Parameters"] is DBNull ? null : (string)dbResult["Parameters"]
                    ));

            return result;

        }

        public CustomActionType(
            Guid id)
        {
            MSPropertyBag dbCustomActionType = _db.GetCustomActionType(id);

            if (dbCustomActionType != null)
            {
                this.id = id;
                name = dbCustomActionType["Name"] is DBNull ? string.Empty : (string)dbCustomActionType["Name"];
                icon = dbCustomActionType["Icon"] is DBNull ? string.Empty : (string)dbCustomActionType["Icon"];
                string parametersString = dbCustomActionType["Parameters"] is DBNull ? string.Empty : (string)dbCustomActionType["Parameters"];
                if (!string.IsNullOrWhiteSpace(parametersString))
                    parameters = JsonSerializer.Deserialize<List<CustomActionTypeParameter>>(parametersString);
                else
                    parameters = new List<CustomActionTypeParameter>();
            }
        }

        public CustomActionType(
            Guid id, string name, string icon, string parameters)
        {
            this.id = id;
            this.name = name;
            this.icon = icon;
            this.parameters = JsonSerializer.Deserialize<List<CustomActionTypeParameter>>(parameters);
        }
    }

    public interface CustomActionBaseConfiguration
    {
        public static CustomActionBaseConfiguration GetInstance(string config)
        {
            return null;
        }
    }

    public class CustomActionHTTPConfiguration: CustomActionBaseConfiguration
    {
        public string uri { get; set; }
 
        public static CustomActionHTTPConfiguration GetInstance(string config)
        {
            if (!string.IsNullOrWhiteSpace(config))
                return JsonSerializer.Deserialize<CustomActionHTTPConfiguration>(config);
            else
                return null;
        }
    }

    public class CustomActionServiceBusConfiguration : CustomActionBaseConfiguration
    {
        public string serviceBusUri { get; set; }
        public string queueName { get; set; }
        public string responseQueueName { get; set; }

        public static CustomActionServiceBusConfiguration GetInstance(string config)
        {
            if (!string.IsNullOrWhiteSpace(config))
                return JsonSerializer.Deserialize<CustomActionServiceBusConfiguration>(config);
            else
                return null;
        }
    }

    public class CustomActionMessage
    {
        public string WorkItemId { get; set; }
        public string WorkItemUrl { get; set; }
        public dynamic Message { get; set; }
    }

    public class CustomActionResponse
    {
        public HttpResponseMessage Response { get; set; }
        public HttpStatusCode StatusCode { get; set; }
        public bool IsSuccessStatusCode { get; set; }
        public string ResponseBody { get; set; }
    }

    public class CustomBaseActionInt
    {
        public Guid id { get; set; }
        public Guid component { get; set; }
        public string name { get; set; }
        public string icon { get; set; }
        public Guid type { get; set; }
        public dynamic configuration { get; set; }
    }

    public class CustomBaseAction
    {
        MSSHNotificationDatabase _db = new MSSHNotificationDatabase();
        protected Guid id;
        protected MSComponent component;
        protected string name;
        protected string icon;
        protected CustomActionType type;
        protected CustomActionBaseConfiguration configuration;

        public Guid Id { get { return id; } }
        public MSComponent Component { get { return component; } }
        public string Name { get { return name; } }
        public string Icon { get { return icon; } }
        public CustomActionType Type { get { return type; } }
        public CustomActionBaseConfiguration Configuration { get { return configuration; } }

        internal static List<CustomBaseAction>GetCustomBaseActionObjectCollection(List<MSPropertyBag> dbResultSet)
        {
            List<CustomBaseAction> result = new List<CustomBaseAction>();

            foreach (MSPropertyBag dbResult in dbResultSet)
            {
                Guid actionId = dbResult["ActionId"] is DBNull ? Guid.Empty : (Guid)dbResult["ActionId"];

                if (actionId != Guid.Empty)
                    result.Add(CustomBaseAction.GetInstance(actionId));
            }

            return result;
        }

        public static List<CustomBaseAction>GetComponentInstances(Guid componentId)
        {
            MSSHNotificationDatabase _db = new MSSHNotificationDatabase();
            List<MSPropertyBag> dbResultSet = _db.GetCustomActionsForComponent(componentId);
            return GetCustomBaseActionObjectCollection(dbResultSet);
        }

        public static List<CustomBaseAction> GetComponentInstances(string internalComponentName)
        {
            MSSHNotificationDatabase _db = new MSSHNotificationDatabase();
            List<MSPropertyBag> dbResultSet = _db.GetCustomActionsForComponent(internalComponentName);
            return GetCustomBaseActionObjectCollection(dbResultSet);
        }

        public static CustomBaseAction GetInstance(Guid actionId)
        {
            MSSHNotificationDatabase _db = new MSSHNotificationDatabase();

            MSPropertyBag actionData = _db.GetCustomAction(actionId);
            if (actionData != null)
            {
                Guid id = actionData["ActionId"] is DBNull ? Guid.Empty : (Guid)actionData["ActionId"];
                string name = actionData["Name"] is DBNull ? string.Empty : (string)actionData["Name"];
                string icon = actionData["Icon"] is DBNull ? string.Empty : (string)actionData["Icon"];
                Guid typeId = actionData["Type"] is DBNull ? Guid.Empty : (Guid)actionData["Type"];
                string config = actionData["Configuration"] is DBNull ? string.Empty : (string)actionData["Configuration"];

                Guid componentId = actionData["ComponentId"] is DBNull ? Guid.Empty : (Guid)actionData["ComponentId"];
                MSComponent component = new MSComponent();
                if (componentId != Guid.Empty)
                {
                    List<MSComponent> components = _db.GetComponent(componentId);
                    if (components != null && components.Count > 0)
                    {
                        component.Id = components[0].Id;
                        component.Name = components[0].Name;
                        component.EntityProperties = components[0].EntityProperties;
                        component.InternalName = components[0].InternalName;
                        component.Capabilities = components[0].Capabilities;
                        component.Icon = components[0].Icon;
                    }
                }

                CustomActionType type = null;
                if (typeId != Guid.Empty)
                    type = new CustomActionType(typeId);

                switch (typeId.ToString().ToLower())
                {
                    case "51ef2a0d-a192-4e91-bc7c-ca08544e5f45":
                        CustomHTTPAction action = new CustomHTTPAction
                        {
                            id = id,
                            name = name,
                            icon = icon,
                            type = type,
                            component = component
                        };
                        action.InitializeConfiguration(config);

                        return action;
                        break;

                    case "da3128f6-65ac-4507-ae53-e7610345391f":
                        CustomServiceBusAction sbAction = new CustomServiceBusAction
                        {
                            id = id,
                            name = name,
                            icon = icon,
                            type = type,
                            component = component
                        };
                        sbAction.InitializeConfiguration(config);

                        return sbAction;
                        break;
                }
            }

            return null;
        }

        internal CustomBaseAction()
        {

        }

        public virtual void InitializeConfiguration(string config)
        {
            configuration = null;
        }

        public virtual CustomActionResponse Run(CustomActionMessage message)
        {
            return null;
        }
    }

    public class CustomHTTPAction: CustomBaseAction
    {
        private static readonly HttpClient client = new HttpClient();

        internal CustomHTTPAction() : base()
        {

        }

        public override void InitializeConfiguration(string config)
        {
            if (!string.IsNullOrWhiteSpace(config))
                configuration = JsonSerializer.Deserialize<CustomActionHTTPConfiguration>(config);
            else
                configuration = null;
        }

        public override CustomActionResponse Run(CustomActionMessage message)
        {
            string uri = ((CustomActionHTTPConfiguration)configuration).uri;
            if (string.IsNullOrWhiteSpace(uri))
                return null;
            else
            {
                CustomActionResponse response = null;
                string serializedMessage = JsonSerializer.Serialize<CustomActionMessage>(message, new JsonSerializerOptions() { MaxDepth = 64, IncludeFields = true });
                using (var content = new StringContent(serializedMessage, System.Text.Encoding.UTF8, "application/json"))
                {
                    HttpResponseMessage result = client.PostAsync(uri, content).Result;
                    response = new CustomActionResponse
                    {
                        Response = result,
                        StatusCode = result.StatusCode,
                        IsSuccessStatusCode = result.IsSuccessStatusCode,
                        ResponseBody = result.Content.ReadAsStringAsync().Result
                    };

                }

                return response;
            }
        }
    }

    public class CustomServiceBusAction : CustomBaseAction
    {
        private static readonly HttpClient client = new HttpClient();

        internal CustomServiceBusAction() : base()
        {

        }

        public override void InitializeConfiguration(string config)
        {
            if (!string.IsNullOrWhiteSpace(config))
                configuration = JsonSerializer.Deserialize<CustomActionServiceBusConfiguration>(config);
            else
                configuration = null;
        }

        public override CustomActionResponse Run(CustomActionMessage message)
        {
            string serviceBusUri = ((CustomActionServiceBusConfiguration)configuration).serviceBusUri;
            string queueName = ((CustomActionServiceBusConfiguration)configuration).queueName;
            if (string.IsNullOrWhiteSpace(serviceBusUri))
                return null;
            else
            {
                CustomActionResponse response = null;
                string serializedMessage = JsonSerializer.Serialize<CustomActionMessage>(message, new JsonSerializerOptions() { MaxDepth = 64, IncludeFields = true });
                TokenCredential clientSecretCredential = new ClientSecretCredential(
                    GlobalConfiguration.Instance.TenantDomain,
                    GlobalConfiguration.Instance.AppId,
                    GlobalConfiguration.Instance.AppSecret);

                try
                {
                    ServiceBusClient client = new ServiceBusClient(
                        serviceBusUri,
                        clientSecretCredential);

                    ServiceBusSender sender = client.CreateSender(queueName);
                    using ServiceBusMessageBatch messageBatch = sender.CreateMessageBatchAsync().Result;
                    if (!messageBatch.TryAddMessage(new ServiceBusMessage(serializedMessage)))
                    {
                        // if it is too large for the batch
                        throw new Exception($"The message is too large to fit in the batch.");
                    }
                    else
                    {

                        try
                        {
                            // Use the producer client to send the batch of messages to the Service Bus queue
                            sender.SendMessagesAsync(messageBatch).Wait();

                            response = new CustomActionResponse
                            {
                                Response = null,
                                StatusCode = HttpStatusCode.OK,
                                IsSuccessStatusCode = true,
                                ResponseBody = "Message successfully posted to Service Bus queue."
                            };
                        }
                        catch (Exception ex)
                        {
                            response = new CustomActionResponse
                            {
                                Response = null,
                                StatusCode = HttpStatusCode.InternalServerError,
                                IsSuccessStatusCode = false,
                                ResponseBody = ex.Message
                            };
                        }
                        finally
                        {
                            // Calling DisposeAsync on client types is required to ensure that network
                            // resources and other unmanaged objects are properly cleaned up.
                            sender.DisposeAsync();
                            client.DisposeAsync();
                        }
                    }
                } catch (Exception ex)
                {
                    response = new CustomActionResponse
                    {
                        Response = null,
                        StatusCode = HttpStatusCode.InternalServerError,
                        IsSuccessStatusCode = false,
                        ResponseBody = ex.Message
                    };
                }

                return response;
            }
        }
    }
}
