using Azure.Core;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace Microsoft.ServiceHealthHub.Core
{
    public class OAuth2Config
    {
        public string ClientId = string.Empty;
        public string ClientSecret = string.Empty;
        public string TenantDomain = string.Empty;
    }

    public sealed class GlobalConfiguration
    {
        private static readonly GlobalConfiguration m_Instance = new GlobalConfiguration();
        private MSSHNotificationDatabase m_db;

        private string m_TenantDomain = string.Empty;
        private string m_AppId = string.Empty;
        private string m_AppSecret = string.Empty;
        private string m_NotificationDatabaseConnectionString = string.Empty;
        private string m_ClientAppId = string.Empty;
        private string m_ClientTenantDomain = string.Empty;
        private bool   m_AzureTranslatorEnabled = false;
        private string m_AzureTranslatorSubscriptionKey = string.Empty;
        private string m_AzureTranslatorResourceLocation = string.Empty;
        private string m_ApplicationInsightsInstrumentationKey = string.Empty;
        private string m_KeyVaultUri = string.Empty;
        private string m_LanguageServiceEndpoint = string.Empty;
        private string m_LanguageServiceKey = string.Empty;
        private bool   m_LanguageServiceEnabled = false;
        private int    m_TextSummarizationSentenceCount = 3;
        private string m_ApplicationTitle = "Service Health Hub";
        private MSApplicationBrandingSettings m_branding = new MSApplicationBrandingSettings();
        private DefaultAzureCredential m_defaultAzureCredential = null;
        private OAuth2Config m_GraphApiAuthConfig = null;
        private OAuth2Config m_CopilotConnectorAuthConfig = null;
        private int    m_MemoryCacheExpiration = 180000; // 3 minutes

        public static GlobalConfiguration Instance
        {
            get
            {
                return m_Instance;
            }
        }

        public string TenantDomain => m_TenantDomain;
        public string ClientAppId => m_ClientAppId;
        public string ClientTenantDomain => m_ClientTenantDomain;
        public string AppId => m_AppId;
        public string AppSecret => m_AppSecret;
        public string NotificationDatabaseConnectionString => m_NotificationDatabaseConnectionString;
        public OAuth2Config GraphApiAuthConfig => m_GraphApiAuthConfig;
        public OAuth2Config CopilotConnectorAuthConfig => m_CopilotConnectorAuthConfig;
        public bool   AzureTranslatorEnabled
        {
            get { return m_AzureTranslatorEnabled; }
            set
            {
                if (value != m_AzureTranslatorEnabled)
                {
                    m_AzureTranslatorEnabled = value;
                    SetConfigurationValue(Constants.c_settings_AzureTranslator_Enabled, value);
                }
            }
        }
        public string AzureTranslatorSubscriptionKey
        {
            get { return m_AzureTranslatorSubscriptionKey; }
            set
            {
                if (m_AzureTranslatorSubscriptionKey != value)
                {
                    m_AzureTranslatorSubscriptionKey = value;
                    SetSecret(Constants.c_settings_AzureTranslator_SubscriptionKey, value);
                }
            }
        }
        public string AzureTranslatorResourceLocation
        {
            get { return m_AzureTranslatorResourceLocation; }
            set
            {
                if (value != m_AzureTranslatorResourceLocation)
                {
                    m_AzureTranslatorResourceLocation = value;
                    SetConfigurationValue(Constants.c_settings_AzureTranslator_ResourceLocation, value);
                }
            }
        }

        public string LanguageServiceEndpoint
        {
            get { return m_LanguageServiceEndpoint; }
            set
            {
                if (value != m_LanguageServiceEndpoint)
                {
                    m_LanguageServiceEndpoint = value;
                    SetConfigurationValue(Constants.c_settings_LanguageService_Endpoint, value);
                }
            }
        }
        public string LanguageServiceSubscriptionKey
        {
            get { return m_LanguageServiceKey; }
            set
            {
                if (m_LanguageServiceKey != value)
                {
                    m_LanguageServiceKey = value;
                    SetSecret(Constants.c_settings_LanguageService_SubscriptionKey, value);
                }
            }
        }
        public bool LanguageServiceEnabled
        {
            get { return m_LanguageServiceEnabled; }
            set
            {
                if (value != m_LanguageServiceEnabled)
                {
                    m_LanguageServiceEnabled = value;
                    SetConfigurationValue(Constants.c_settings_LanguageService_Enabled, value);
                }
            }
        }
        public int TextSummarizationSentenceCount
        {
            get { return m_TextSummarizationSentenceCount; }
            set
            {
                if (value != m_TextSummarizationSentenceCount)
                {
                    m_TextSummarizationSentenceCount = value;
                    SetConfigurationValue(Constants.c_settings_TextSummarization_SentenceCount, value);
                }
            }
        }
        public string ApplicationInsightsInstrumentationKey => m_ApplicationInsightsInstrumentationKey;
        public string KeyVaultUri => m_KeyVaultUri;

        public string ApplicationTitle => m_ApplicationTitle;

        public MSApplicationBrandingSettings Branding
        {
            get { return m_branding; }
            set
            {
                if (value != m_branding)
                {
                    m_branding = value;
                    SetConfigurationValue(Constants.c_settings_Branding, value);
                }
            }
        }

        public DefaultAzureCredential DefaultAzureCredential
        {
            get
            {
                SqlConnectionStringBuilder sqlConnString = new SqlConnectionStringBuilder(NotificationDatabaseConnectionString);
                if (string.IsNullOrEmpty(sqlConnString.UserID) && m_defaultAzureCredential == null)
                {
                    m_defaultAzureCredential = new DefaultAzureCredential();
                }

                return m_defaultAzureCredential;
            }
        }

        public int MemoryCacheExpiration => m_MemoryCacheExpiration;

        // Explicit static constructor to tell C# compiler
        // not to mark type as beforefieldinit
        static GlobalConfiguration()
        {
        }

        private GlobalConfiguration()
        {

        }

        public string GetConfigurationValue(string name)
        {
            object value = m_db.GetConfigurationValue(name);
            return value?.ToString();
        }

        public void SetConfigurationValue(string name, object value)
        {
            m_db.SetConfigurationValue(name, value);
        }

        public string GetSecret(string name)
        {
            if (string.IsNullOrWhiteSpace(KeyVaultUri))
                throw new Exception("Key Vault Uri is not specified in the configuration.");

            SecretClientOptions options = new()
            {
                Retry =
                {
                    Delay= TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = RetryMode.Exponential
                 }
            };

            ClientSecretCredentialOptions o = new ClientSecretCredentialOptions();
            o.AdditionallyAllowedTenants.Add("*");

            ClientSecretCredential clientSecretCredential = new(TenantDomain, AppId,AppSecret, o);
            var client = new SecretClient(new Uri(KeyVaultUri), clientSecretCredential, options);

            try
            {
                KeyVaultSecret secret = client.GetSecret(name);
                return secret.Value;
            } 
            catch
            {
                return string.Empty;
            }
        }

        public void SetSecret(string name, string value)
        {
            if (string.IsNullOrWhiteSpace(KeyVaultUri))
                throw new Exception("Key Vault Uri is not specified in the configuration.");

            SecretClientOptions options = new()
            {
                Retry =
                {
                    Delay= TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = RetryMode.Exponential
                 }
            };

            ClientSecretCredential clientSecretCredential = new(TenantDomain, AppId, AppSecret);
            var client = new SecretClient(new Uri(KeyVaultUri), clientSecretCredential, options);

            KeyVaultSecret secret = new KeyVaultSecret(name, value);
            client.SetSecret(secret);
        }

        public static void LoadConfiguration(IConfiguration config)
        {
            Instance.m_TenantDomain = config["TenantDomain"];
            Instance.m_ClientAppId = config["ClientAppId"];
            Instance.m_ClientTenantDomain = config["ClientTenantDomain"];
            Instance.m_AppId = config["AppId"];
            Instance.m_AppSecret = config["AppSecret"];
            Instance.m_KeyVaultUri = config["KeyVaultUri"];
            Instance.m_AzureTranslatorSubscriptionKey = Instance.GetSecret(Constants.c_settings_AzureTranslator_SubscriptionKey);
            Instance.m_NotificationDatabaseConnectionString = Instance.GetSecret("ConnectionString");
            Instance.m_db = new MSSHNotificationDatabase();

            string graphApiAuthConfigKey = config["GraphApiAuthConfig"];
            if (!string.IsNullOrEmpty(graphApiAuthConfigKey))
            {
                try
                {
                    string graphApiAuthConfigJson = Instance.GetSecret(graphApiAuthConfigKey);
                    Instance.m_GraphApiAuthConfig = JsonConvert.DeserializeObject<OAuth2Config>(graphApiAuthConfigJson);
                } catch
                {

                }
            }

            string copilotConnectorConfigKey = config["CopilotConnectorAuthConfig"];
            if (!string.IsNullOrEmpty(copilotConnectorConfigKey))
            {
                try
                {
                    string copilotConnectorAuthConfigJson = Instance.GetSecret(copilotConnectorConfigKey);
                    Instance.m_CopilotConnectorAuthConfig = JsonConvert.DeserializeObject<OAuth2Config>(copilotConnectorAuthConfigJson);
                }
                catch
                {

                }
            }

            bool.TryParse(Instance.GetConfigurationValue(Constants.c_settings_AzureTranslator_Enabled), out Instance.m_AzureTranslatorEnabled);
            Instance.m_AzureTranslatorResourceLocation = Instance.GetConfigurationValue(Constants.c_settings_AzureTranslator_ResourceLocation);
            Instance.m_ApplicationInsightsInstrumentationKey = config["ApplicationInsightsInstrumentationKey"];
            Instance.m_ApplicationTitle = config["AppTitle"] != null && !string.IsNullOrWhiteSpace(config["AppTitle"]) ? config["AppTitle"] : "Service Health Hub";

            string brandingStr = Instance.GetConfigurationValue(Constants.c_settings_Branding);
            if (!string.IsNullOrWhiteSpace(brandingStr))
                try
                {
                    Instance.m_branding = JsonConvert.DeserializeObject<MSApplicationBrandingSettings>(brandingStr);
                }
                catch
                {
                    Instance.m_branding = new MSApplicationBrandingSettings();
                }

            bool.TryParse(Instance.GetConfigurationValue(Constants.c_settings_LanguageService_Enabled), out Instance.m_LanguageServiceEnabled);
            Instance.m_LanguageServiceEndpoint = Instance.GetConfigurationValue(Constants.c_settings_LanguageService_Endpoint);
            
            Instance.m_LanguageServiceKey = Instance.GetSecret(Constants.c_settings_LanguageService_SubscriptionKey);
            if (!int.TryParse(Instance.GetConfigurationValue(Constants.c_settings_TextSummarization_SentenceCount), out Instance.m_TextSummarizationSentenceCount))
                Instance.m_TextSummarizationSentenceCount = 3;
        }

    }
}
