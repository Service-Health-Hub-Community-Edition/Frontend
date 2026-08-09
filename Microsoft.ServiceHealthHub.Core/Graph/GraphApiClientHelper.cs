using Azure.Identity;
using Microsoft.Graph;
using Microsoft.ServiceHealthHub.Core;

namespace Microsoft.ServiceHealthHub.Graph
{
    public static class GraphApiClientHelper
    {
        private static string[] scopes = { "https://graph.microsoft.com/.default" };

        private static ClientSecretCredential clientSecretCredential = GlobalConfiguration.Instance.GraphApiAuthConfig != null ? 
            new ClientSecretCredential(
            GlobalConfiguration.Instance.GraphApiAuthConfig.TenantDomain,
            GlobalConfiguration.Instance.GraphApiAuthConfig.ClientId,
            GlobalConfiguration.Instance.GraphApiAuthConfig.ClientSecret) :
            new ClientSecretCredential(
            GlobalConfiguration.Instance.TenantDomain,
            GlobalConfiguration.Instance.AppId,
            GlobalConfiguration.Instance.AppSecret);

        public static GraphServiceClient Client = new GraphServiceClient(clientSecretCredential, scopes);
    }

    public static class CopilotConnectorClientHelper
    {
        private static string[] scopes = { "https://graph.microsoft.com/.default" };

        private static ClientSecretCredential clientSecretCredential = GlobalConfiguration.Instance.CopilotConnectorAuthConfig != null ?
            new ClientSecretCredential(
            GlobalConfiguration.Instance.CopilotConnectorAuthConfig.TenantDomain,
            GlobalConfiguration.Instance.CopilotConnectorAuthConfig.ClientId,
            GlobalConfiguration.Instance.CopilotConnectorAuthConfig.ClientSecret) :
            new ClientSecretCredential(
            GlobalConfiguration.Instance.TenantDomain,
            GlobalConfiguration.Instance.AppId,
            GlobalConfiguration.Instance.AppSecret);

        public static GraphServiceClient Client = new GraphServiceClient(clientSecretCredential, scopes);
    }
}
