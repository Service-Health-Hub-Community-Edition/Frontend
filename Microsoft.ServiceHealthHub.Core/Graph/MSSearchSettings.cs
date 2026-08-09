using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.ServiceHealthHub.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.ServiceHealthHub.Core.Graph
{
    public class MSNameValuePair
    {
        public string name { get; set; }
        public object value { get; set; }
    }

    public class MSNameValuePairList : System.Collections.Generic.List<MSNameValuePair>
    {
        public MSNameValuePairList() { }
        public object GetValue(string name)
        {
            return this.Where((n) => n.name == name).First().value;
        }

        public void SetValue(string name, object value)
        {
            var o = this.Where((n) => n.name == name).First();
            if (o != null)
            {
                o.value = value;
            }
            else
            {
                Add(new MSNameValuePair { name = name, value = value });
            }
        }
    }

    public enum SearchOperation
    {
        create,
        configure,
        enable,
        disable,
        schemaUpdate
    }

    public class SearchConfigRequest
    {
        public SearchOperation? operation = null;
        public string rootUrl = null;
    }

    public class MSSearchConnectorDetails
    {
        public bool SchemaUpdateAvailable = false;
        public string ConnectorStatus = string.Empty;
    }

    public class MSSearchConnectorSettings
    {
        public bool Configured = false;
        public bool Enabled = false;
        public DateTime SchemaUpdated;
        public string RootUrl = string.Empty;
        public object SyncACL = null;

        private MSConnector connector = null;

        public MSSearchConnectorSettings()
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            connector = db.GetConnector(new System.Guid("a6a162f4-4910-4c64-a5fd-18a01665c88d"));
            if (connector != null)
            {
                MSNameValuePairList configuration = ((Newtonsoft.Json.Linq.JArray)(connector.Configuration)).ToObject<MSNameValuePairList>();
                Configured = (bool)configuration.GetValue("Configured");
                Enabled = (bool)configuration.GetValue("Enabled");
                SchemaUpdated = (DateTime)configuration.GetValue("SchemaUpdated");
                RootUrl = (string)configuration.GetValue("RootUrl");
                SyncACL = (object)configuration.GetValue("SyncACL");
            }
        }

        public void Connect(ExternalConnection externalConnection)
        {
            if (externalConnection != null && connector != null)
            {
                MSSHNotificationDatabase db = new MSSHNotificationDatabase();

                MSNameValuePairList configuration = ((Newtonsoft.Json.Linq.JArray)(connector.Configuration)).ToObject<MSNameValuePairList>();
                Configured = true;
                Enabled = true;
                configuration.SetValue("Configured", Configured);
                configuration.SetValue("Enabled", Enabled);

                connector.Configuration = Newtonsoft.Json.Linq.JArray.FromObject(configuration);
                db.UpdateConnector(connector);
            }
        }

        public void Enable()
        {
            if (connector != null)
            {
                MSSHNotificationDatabase db = new MSSHNotificationDatabase();

                Enabled = true;
                MSNameValuePairList configuration = ((Newtonsoft.Json.Linq.JArray)(connector.Configuration)).ToObject<MSNameValuePairList>();
                configuration.SetValue("Enabled", Enabled);
                connector.Configuration = Newtonsoft.Json.Linq.JArray.FromObject(configuration);

                db.UpdateConnector(connector);
            }
        }

        public void Disable()
        {
            if (connector != null)
            {
                MSSHNotificationDatabase db = new MSSHNotificationDatabase();

                Enabled = false;
                MSNameValuePairList configuration = ((Newtonsoft.Json.Linq.JArray)(connector.Configuration)).ToObject<MSNameValuePairList>();
                configuration.SetValue("Enabled", Enabled);
                connector.Configuration = Newtonsoft.Json.Linq.JArray.FromObject(configuration);
                db.UpdateConnector(connector);
            }
        }

        public void SetSchemaTimestamp(DateTime timestamp)
        {
            if (connector != null)
            {
                MSSHNotificationDatabase db = new MSSHNotificationDatabase();

                SchemaUpdated = timestamp;
                MSNameValuePairList configuration = ((Newtonsoft.Json.Linq.JArray)(connector.Configuration)).ToObject<MSNameValuePairList>();
                configuration.SetValue("SchemaUpdated", timestamp);
                connector.Configuration = Newtonsoft.Json.Linq.JArray.FromObject(configuration);
                db.UpdateConnector(connector);
            }
        }

        public void SetConfig(SearchConfigRequest config)
        {
            if (connector != null)
            {
                MSSHNotificationDatabase db = new MSSHNotificationDatabase();

                if (config.rootUrl != null)
                {
                    RootUrl = config.rootUrl;
                    MSNameValuePairList configuration = ((Newtonsoft.Json.Linq.JArray)(connector.Configuration)).ToObject<MSNameValuePairList>();
                    configuration.SetValue("RootUrl", RootUrl);
                    connector.Configuration = Newtonsoft.Json.Linq.JArray.FromObject(configuration);
                    db.UpdateConnector(connector);
                }
            }
        }
    }



    public class MSSearchSettings
    {
        public MSSearchConnectorSettings ConnectorSettings = new MSSearchConnectorSettings();
        public MSSearchConnectorDetails Details = new MSSearchConnectorDetails();
        public ExternalConnection GraphAPIConnector { get; set; }

        public MSSearchSettings()
        {
            try
            {
                var connectionsRequest = CopilotConnectorClientHelper.Client.External.Connections["mscommshub"];

                Task<ExternalConnection> t = Task.Run(async () =>
                {
                    ExternalConnection res = await connectionsRequest.GetAsync();
                    return res;
                });

                t.Wait();

                GraphAPIConnector = t.Result;
                Details.ConnectorStatus = GraphAPIConnector.State.ToString();
                Details.SchemaUpdateAvailable = SchemaUpdateAvailable();
            }
            catch
            {
                GraphAPIConnector = null;

                Details.ConnectorStatus = "Not found";
                Details.SchemaUpdateAvailable = false;
                ConnectorSettings.Configured = false;
            }


        }

        public async Task<ExternalConnection> CreateConnection()
        {
            ExternalConnection connection = new ExternalConnection()
            {
                Id = "mscommshub",
                Description = "Connection to index Service Health Hub, data from Message Center, Service Health, Microsoft 365 Roadmap, Release Planner, Azure Service Health and Azure Updates",
                Name = "Service Health Hub"
            }; ;

            var conn = await CopilotConnectorClientHelper.Client.External.Connections.PostAsync(connection);

            return conn;
        }

        public bool SchemaUpdateAvailable()
        {
            var db = new MSSHNotificationDatabase();
            var connectorTemplates = db.GetConnectorDefinitionTemplates(new System.Guid("4d7bed24-698e-47ce-ae13-44a15982906f"), "schema");

            if (connectorTemplates != null && connectorTemplates.Count > 0)
            {
                MSConnectorDefinitionTemplate connectorTemplate = connectorTemplates[0];

                return ConnectorSettings != null && ConnectorSettings.SchemaUpdated < connectorTemplate.Modified;
            }
            else
            {
                throw new System.Exception("Search schema is not registered within the database. Please run the database upgrade procedure first and repeat the operation");
            }
        }

        public async Task<Schema> ConfigureSchema(ExternalConnection connection)
        {
            Schema returnSchema = null;

            var db = new MSSHNotificationDatabase();
            var connectorTemplates = db.GetConnectorDefinitionTemplates(new System.Guid("4d7bed24-698e-47ce-ae13-44a15982906f"), "schema");

            if (connectorTemplates != null && connectorTemplates.Count > 0)
            {
                MSConnectorDefinitionTemplate connectorTemplate = connectorTemplates[0];
                Schema schema = JsonConvert.DeserializeObject<Schema>(connectorTemplate.Template);

                returnSchema = await CopilotConnectorClientHelper.Client.External.Connections[connection.Id].Schema.PatchAsync(schema)
                    .ContinueWith(t => {
                        ConnectorSettings.SetSchemaTimestamp(connectorTemplate.Modified);
                        return t.Result;
                    }

                    );
            }

            return returnSchema;
        }

        public async void Configure()
        {
            if (GraphAPIConnector == null)
            {
                await CreateConnection().ContinueWith(async t => {
                    GraphAPIConnector = t.Result;
                    await ConfigureSchema(GraphAPIConnector);
                }
                ).ContinueWith(t => ConnectorSettings.Connect(GraphAPIConnector));
            }
            else
            {
                if (ConnectorSettings != null)
                    ConnectorSettings.Connect(GraphAPIConnector);
            }
        }
    }
}
