using Microsoft.Graph.Education.Classes.Item.Assignments.Item.Submissions.Item.Return;
using Microsoft.Graph.External.Connections.Item.Items.Item.MicrosoftGraphExternalConnectorsAddActivities;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.ServiceHealthHub.Azure;
using Microsoft.ServiceHealthHub.Core.Graph;
using Microsoft.ServiceHealthHub.Graph;
using Newtonsoft.Json;
using System.Data.SqlTypes;
using System.Dynamic;
using System.Security.Cryptography;
using System.Text;

namespace Microsoft.ServiceHealthHub.Core
{
    public class MSServiceNotification
    {
        internal string m_Id = string.Empty;
        internal DateTime m_LastUpdatedTime = SqlDateTime.MinValue.Value;
        internal string m_Data = string.Empty;
        internal bool m_public = false;
        internal string m_comments = string.Empty;
        internal string m_WorkItemID = null;
        internal string m_WorkItemURL = null;
        internal dynamic m_ExtendedProperties = null;
        internal string m_State = null;

        public string Id => m_Id;
        public DateTime LastUpdatedTime => m_LastUpdatedTime;
        public string Data => m_Data;
        public bool Public => m_public;
        public string Comments => m_comments;
        public string WorkItemID => m_WorkItemID;
        public string WorkItemURL => m_WorkItemURL;
        public dynamic ExtendedProperties => m_ExtendedProperties;
        public string State => m_State;

        public MSServiceNotification()
        {

        }

        public static MSServiceNotification GetNotification(string id)
        {
            MSServiceNotification res = new MSServiceNotification();

            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.GetNotification(id, res);

            return res;
        }
    }

    public class MSServiceIncident
    {
        internal string m_Id = string.Empty;
        internal string m_Title = string.Empty;
        internal string m_Workload = string.Empty;
        internal DateTime m_StartTime = SqlDateTime.MinValue.Value;
        internal DateTime m_EndTime = SqlDateTime.MinValue.Value;

        public string Id => m_Id;
        public string Title => m_Title;
        public string Workload => m_Workload;
        public DateTime StartTime => m_StartTime;
        public DateTime EndTime => m_EndTime;

        public MSServiceIncident()
        {

        }

        public static List<MSServiceIncident> GetIncidents(string workloads, string startDate, string endDate)
        {
            List<MSServiceIncident> res = new List<MSServiceIncident>();

            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.GetIncidents(workloads, startDate, endDate, res);

            return res;
        }
    }

    public class MSServiceWorkloads
    {
        public MSServiceWorkloads()
        {

        }

        public static List<string> GetServiceWorkloads()
        {
            List<string> res = new List<string>();

            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.GetServiceWorkloads(res);

            return res;
        }
    }

    public class MSRoadmapNotificationCollection: List<MSRoadmapNotification>
    { }

    public class MSRoadmapNotification
    {
        internal string m_Id = string.Empty;
        internal string m_Title = string.Empty;
        internal string m_Status = string.Empty;       
        internal DateTime m_Published = SqlDateTime.MinValue.Value;
        internal DateTime m_LastUpdated = SqlDateTime.MinValue.Value;
        internal object m_Category = null;
        internal object m_Tags = null;
        internal object m_CloudInstances = null;
        internal object m_Products = null;
        internal object m_ReleasePhase = null;
        internal object m_Platforms = null;
        internal string m_Link = string.Empty;
        internal string m_MoreInfoLink = string.Empty;
        internal string m_Description = string.Empty;
        internal string m_AvailabilityDate = string.Empty;
        internal DateTime m_AvailabilityFrom = SqlDateTime.MinValue.Value;
        internal DateTime m_AvailabilityTo = SqlDateTime.MinValue.Value;
        internal string m_PublicPreviewDate = string.Empty;
        internal DateTime m_PublicPreviewFrom = SqlDateTime.MinValue.Value;
        internal DateTime m_PublicPreviewTo = SqlDateTime.MinValue.Value;
        internal string m_WorkItemID = string.Empty;
        internal string m_WorkItemURL = string.Empty;
        internal dynamic m_ExtendedProperties = null;

        public string Id => m_Id;
        public string Title => m_Title;
        public string Status => m_Status;
        public DateTime Published => m_Published;
        public DateTime LastUpdated => m_LastUpdated;
        public object Category => m_Category;
        public object Products => m_Products;
        public object Tags => m_Tags;
        public object CloudInstances => m_CloudInstances;
        public object ReleasePhase => m_ReleasePhase;
        public object Platforms => m_Platforms;
        public string Link => m_Link;
        public string MoreInfoLink => m_MoreInfoLink;

        public string Description => m_Description;
        public string AvailabilityDate => m_AvailabilityDate;
        public DateTime AvailabilityFrom => m_AvailabilityFrom;
        public DateTime AvailabilityTo => m_AvailabilityTo;
        public string PublicPreviewDate => m_PublicPreviewDate;
        public DateTime PublicPreviewFrom => m_PublicPreviewFrom;
        public DateTime PublicPreviewTo => m_PublicPreviewTo;
        public string WorkItemID => m_WorkItemID;
        public string WorkItemURL => m_WorkItemURL;
        public dynamic ExtendedProperties => m_ExtendedProperties;

        public MSRoadmapNotification()
        {

        }


        public static MSRoadmapNotification GetNotification(string id)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            return db.GetRoadmapNotification(id);
        }

        public static MSRoadmapNotificationCollection GetNotifications()
        {
            MSRoadmapNotificationCollection res = new MSRoadmapNotificationCollection();

            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.GetRoadmapNotificationCollection(res);

            return res;
        }
    }

    public class MSPublicIncidentCollection : List<MSPublicIncident>
    { }

    public class MSPublicIncident
    {
        internal string _id = string.Empty;
        internal string _title = string.Empty;
        internal string _service = string.Empty;
        internal DateTime _startTime;
        internal DateTime _lastModified;
        internal string _comments = string.Empty;

        public string Id => _id;
        public string Title => _title;
        public string Service => _service;
        public DateTime StartTime => _startTime;
        public DateTime LastModified => _lastModified;
        public string Comments => _comments;

        public MSPublicIncident()
        {

        }

        public void SetPublicIncident(string title, string comments)
        {
            _title = title;
            _comments = comments;
        }

        public static MSPublicIncidentCollection GetPublicIncidents()
        {
            MSPublicIncidentCollection res = new MSPublicIncidentCollection();

            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.GetPublicIncidentCollection(res);

            return res;
        }
    }

    public class MSPublicMessageCollection : List<MSPublicMessage>
    { }

    public class MSPublicMessageServicesList
    {
        public List<string> services;
    }

    public class MSPublicMessage
    {
        internal string _id = string.Empty;
        internal string _title = string.Empty;
        internal List<string> _services = new List<string>();
        internal DateTime _startTime;
        internal DateTime _lastModified;
        internal string _content = string.Empty;
        internal string _comments = string.Empty;

        public string Id => _id;
        public string Title => _title;
        public List<string> Services => _services;
        public DateTime StartTime => _startTime;
        public DateTime LastModified => _lastModified;
        public string Content => _content;
        public string Comments => _comments;

        public MSPublicMessage()
        {

        }

        public MSPublicMessage(string id)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.GetPublicMessage(id, this);
        }

        public void SetPublicMessage(string title, string comments)
        {
            _title = title;
            _comments = comments;
        }

        public void SetPublicMessage(string title, string content, string comments)
        {
            _title = title;
            _content = content;
            _comments = comments;
        }

        public static void Publish(string id, string type, string comments, MSUserInfo userInfo)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.PublishMessage(id, type, comments, userInfo);
            Cache.Instance.ReloadPersistedContent(id, type);
        }

        public static void Unpublish(string id, string type, MSUserInfo userInfo)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.UnpublishMessage(id, type, userInfo);
            Cache.Instance.ReloadPersistedContent(id, type);
        }

        public static MSPublicMessage GetPublicMessage(string id, string lang = "")
        {
            MSPublicMessage publicMessage = new MSPublicMessage(id);

            if (!string.IsNullOrWhiteSpace(lang))
            {
                string[] messages = new string[] { publicMessage.Title, publicMessage.Comments, publicMessage.Content };

                List<MSTranslationCollection> result = MSAzureTranslator.Translate(messages, lang, true);

                MSTranslationCollection titleTranslation = result.Find(translation => translation.OriginalMessage == publicMessage.Title);
                MSTranslationCollection commentsTranslation = result.Find(translation => translation.OriginalMessage == publicMessage.Comments);
                MSTranslationCollection contentTranslation = result.Find(translation => translation.OriginalMessage == publicMessage.Content);

                if (!string.IsNullOrWhiteSpace(titleTranslation?.Translations?[lang]?.Message))
                {
                    publicMessage.SetPublicMessage(
                        titleTranslation?.Translations?[lang]?.Message,
                        contentTranslation?.Translations?[lang]?.Message,
                        commentsTranslation?.Translations?[lang]?.Message);
                }
            }

            return publicMessage;
        }

        public static MSPublicMessageCollection GetPublicMessages()
        {
            MSPublicMessageCollection res = new MSPublicMessageCollection();

            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.GetPublicMessageCollection(res);

            return res;
        }
    }

    public class MSTranslationCollection
    {
        public string OriginalMessage = string.Empty;
        public Dictionary<string, MSTranslation> Translations = new Dictionary<string, MSTranslation>();
    }

    public class MSTranslation
    {
        internal string _hash = string.Empty;
        internal string _language = string.Empty;
        internal string _message = string.Empty;

        public string Hash
        {
            get
            {
                return _hash;
            }
            set
            {
                _hash = value;
            }
        }

        public string Language
        {
            get
            {
                return _language;
            }
            set
            {
                _language = value;
            }
        }
        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                _message = value;
            }
        }
    }

    public class MSSyncConfigEntry
    {
        internal Int32 _id = -1;
        internal string _component = string.Empty;
        internal string _element = string.Empty;
        internal string _config = string.Empty;
        internal DateTime _created;
        internal string _createdBy = string.Empty;
        internal DateTime _modified;
        internal string _modifiedBy = string.Empty;

        public Int32 Id => _id;
        public string Component => _component;
        public string Element => _element;
        public string Config => _config;
        public DateTime Created => _created;
        public string CreatedBy => _createdBy;
        public DateTime Modified => _modified;
        public string ModifiedBy => _modifiedBy;

        public MSSyncConfigEntry()
        {

        }

        public MSSyncConfigEntry(string component, string element)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.GetSyncConfigEntry(component, element, this);
        }

        public void SetConfigEntry(string component, string element, string config)
        {
            _component = component;
            _element = element;
            _config = config;
        }

        public void Update(MSUserInfo userInfo)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            db.UpdateSyncConfigEntry(this, userInfo);
        }
    }

    public class MSLicenseStatistics
    {
        public string DisplayName { get; set; }
        public string SkuPartNumber { get; set; }
        public int Enabled { get; set; }
        public int ConsumedUnits { get; set; }
        public int AvailableUnits { get; set; }
        public int Suspended { get; set; }
        public int Warning { get; set; }
        public DateTime Date { get; set; }
    }

    public class MSLicenseStatisticsExt
    {
        public string DisplayName { get; set; }
        public string SkuPartNumber { get; set; }
        public int Enabled { get; set; }
        public int ConsumedUnits { get; set; }
        public int Suspended { get; set; }
        public int Warning { get; set; }
        public DateTime Timestamp { get; set; }
    }

    public class MSLicenseForecast
    {
        public string SkuPartNumber { get; set; }
        public int NewHires { get; set; }
        public int Leavers { get; set; }
        public int Balance { get; set; }
        public DateTime Month { get; set; }
    }

    public class MSConnectorConfigurationEntry
    {
        public string name { get; set; }
        public object value { get; set; }
    }

    public class MSConnectorParameterDefinitionEntry
    {
        public int id { get; set; }
        public string name { get; set; }
        public string displayName { get; set; }
        public string type { get; set; }
        public string description { get; set; }
        public string parameterType { get; set; }
    }

    public class MSConnector
    {
        public Guid ConnectorId { get; set; }
        public string Name { get; set; }
        public Guid Type { get; set; }
        public object Configuration { get; set; }
        public string ConnectorTypeName { get; set; }
        public string ConnectorType { get; set; }
        public string Icon { get;  set; }
        public bool System { get; set; }
        public bool Hidden { get; set; }
        public object ParameterDefinition { get; set; }
    }

    public class MSConnectorDefinition
    {
        public Guid ConnectorId { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string Icon { get; set; }
        public bool Unique { get; set; }
        public bool System { get; set; }
        public object Parameters { get; set; }
        public bool Hidden { get; set; }
    }

    public class MSConnectorDefinitionTemplate
    {
        public Guid ConnectorDefinition { get; set; }
        public string Entity { get; set; }
        public string Type { get; set; }
        public string Template { get; set; }
        public DateTime Created { get; set; }
        public string CreatedBy { get; set; }
        public DateTime Modified { get; set; }
        public string ModifiedBy { get; set; }
    }

    public class MSRoute
    {
        public Guid Id { get; set; }
        public int Order { get; set; }
        public string Name { get; set; }
        public string Icon { get; set; }
        public string? Language { get; set; }
        public bool StopProcessingOnMatch { get; set; }
        public bool HideWorkItemLink { get; set; }
        public object Conditions { get; set; }
        public MSConnector Connector { get; set; }
        public object ConnectorConfiguration { get; set; }
        public Guid Component { get;  set; }

    }

    public class MSRouteId
    {
        public Guid Id { get; set; }

        public MSRouteInt GetRouteObject()
        {
            MSRouteInt res = new MSRouteInt();
            res.Id = Id;

            return res;
        }
    }

    public class MSRouteInt
    {
        public Guid Id { get; set; }
        public int Order { get; set; }
        public string Name { get; set; }
        public string Icon { get; set; }
        public string? Language { get; set; }
        public bool StopProcessingOnMatch { get; set; }
        public bool HideWorkItemLink { get; set; }
        public object Conditions { get; set; }
        public Guid Connector { get; set; }
        public object ConnectorConfiguration { get; set; }
        public Guid Component { get; set; }

    }

    public class MSComponent
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string InternalName { get; set; }
        public string? Icon { get; set; }
        public List<string> Capabilities { get; set; }
        public object EntityProperties { get; set; }
    }

    public class MSMonthlyActiveUsers
    {
        public string Service { get; set; }
        public int Active { get; set; }
        public int Inactive { get; set; }
        public DateTime ReportDate { get; set; }
    }

    public class MSSummaryCacheEntry
    {
        public string id { get; set; }
        public DateTime? timestamp { get; set; }
        public object contents { get; set; }
    }

    public class MSSHNotificationDatabase
    {
        private const string sql_GetConfigurationValue = "SELECT [SerializedValue] FROM [dbo].[Config] WHERE [Key]=@Name";
        private const string sql_GetNotification = "SELECT * FROM [dbo].[vw_GetCommunications] WHERE [ID]=@ID";
        private const string sql_GetNotifications = "SELECT * FROM [dbo].[vw_GetCommunications]";
        private const string sql_IncidentTimeline = "SELECT [ID], [Title], [Workload], [StartTime], [EndTime] FROM [dbo].[AllComms]";
        private const string sql_GetServiceWorkloads = "SELECT [Workload] FROM [dbo].[AllServiceHealthWorkloads] ORDER BY [Workload]";
        private const string sql_GetFutureRoadmapItems = "SELECT * FROM [dbo].[AllRoadmapCommsDetailed] WHERE[AvailabilityFrom] >= DATEADD(month, DATEDIFF(month, 0, GETUTCDATE()), 0) ORDER BY[AvailabilityFrom]";
        private const string sql_GetRoadmapItem = "SELECT * FROM [dbo].[AllRoadmapCommsDetailed] WHERE [ID] = @ID";
        private const string sql_GetRoadmapStatistics = "SELECT [inDevelopment], [rollingOut] FROM [dbo].[vw_RoadmapStatistics]";
        private const string sql_GetPublicIncidents = "SELECT * FROM [dbo].[AllPublicIncidents]";
        private const string sql_GetPublicMessages = "SELECT [ID], [Title], [Services], [StartTime], [EndTime], [lastModifiedDateTime], [Comments] FROM [dbo].[AllPublicMessages]";
        private const string sql_GetPublicMessage = "SELECT * FROM [dbo].[AllPublicMessages] WHERE [ID]=@ID";
        private const string sql_GetTranslationFromCache = "SELECT * FROM [dbo].[TranslationCache] WHERE [Hash]=@Hash AND [Language]=@Language";
        private const string sql_GetTranslationsFromCache = "SELECT * FROM [dbo].[TranslationCache] WHERE [Hash] IN ({0}) AND [Language]=@Language";
        private const string sql_GetSyncConfigEntry = "SELECT * FROM [dbo].[SyncConfig] WHERE [Component]=@Component AND [Element]=@Element";
        private const string sql_GetAllLicenseStatistics = "SELECT * FROM [dbo].[AllLicenseStatistics] WHERE [Timestamp] >= DATEADD(month, -12, GETUTCDATE())";
        private const string sql_GetLastLicenseStatistics = "SELECT * FROM [dbo].[LastLicenseStatistics]";
        private const string sql_GetAllLicensesForecast = "SELECT * FROM [dbo].[vw_GetAllLicenseForecast]";
        private const string sql_GetConnectors = "SELECT * FROM [dbo].[vw_GetConnectors]";
        private const string sql_GetConnectorDefinitions = "SELECT * FROM [dbo].[ConnectorDefinition]";
        private const string sql_GetConnectorDefinitionTemplates = "SELECT * FROM [dbo].[ConnectorDefinitionTemplates] WHERE [ConnectorDefinition]=@ConnectorDefinition";
        private const string sql_GetRoutes = "SELECT * FROM [dbo].[vw_GetRoutes]";
        private const string sql_GetComponents = "SELECT * FROM [dbo].[Components]";
        private const string sql_GetMAU = "SELECT * FROM [dbo].[GetLatestMonthlyUsageReport] ORDER BY [Service]";
        private const string sql_GetCustomActionsForComponent = "SELECT * FROM [dbo].[CustomActions] WHERE [ComponentId]=@ComponentId";
        private const string sql_GetCustomAction = "SELECT * FROM [dbo].[CustomActions] WHERE [ActionId]=@ActionId";
        private const string sql_GetCustomActionType = "SELECT * FROM [dbo].[CustomActionTypes] WHERE [ActionTypeId]=@ActionTypeId";
        private const string sql_GetCustomActionTypes = "SELECT * FROM [dbo].[CustomActionTypes]";
        private const string sql_GetImages = "SELECT * FROM [dbo].[Images]";
        private const string sql_GetAuditLog = "SELECT * FROM [dbo].[ActivityLog]";
        private const string sql_GetUserProfile = "SELECT * FROM [dbo].[UserProfiles] WHERE [ObjectId]=@ObjectId";
        private const string sql_GetJobStatistics = @"
SELECT
    [js].*,
    (SELECT TOP (1) [Activity] FROM [dbo].[ActivityLog] WHERE [CorrelationID]=js.CorrelationId AND ([Activity]='JobCompleted' OR [Activity]='JobFailed')) AS JobState
FROM
    [dbo].[JobStatistics] AS [js]
INNER JOIN 
    [dbo].[Components] AS [c] ON ([js].[JobName]=[c].[InternalName])
WHERE 
    [c].[ComponentId]=@JobId";
        private const string sql_GetTagDefinitions = "SELECT [Id], [TagId], [Name], [Type], [ItemCount], [LastUsed] FROM [dbo].[vw_GetTagDefinitions]";
        private const string sql_GetTagDefinition = "SELECT [Id], [TagId], [Name], [Type], [ItemCount], [LastUsed] FROM [dbo].[vw_GetTagDefinitions] WHERE [TagId] = @TagId";
        private const string sql_GetTags = "SELECT [Id], [MessageId], [Type], [TagId], [Modified] FROM [dbo].[Tags] WHERE [MessageId]=@MessageId AND [Type]=@Type";
        private const string sql_GetTagCollection = "SELECT * FROM [dbo].[fn_GetTagCollection](@CommunicationIds)";
        private const string sql_GetJobStatisticsForComm = @"
SELECT 
    [js].* 
FROM 
    [dbo].[JobStatistics] AS [js] 
INNER JOIN 
    [dbo].[Components] AS [c] ON ([js].[JobName]=[c].[InternalName])
WHERE 
    [js].[CorrelationId] IN (
        SELECT
            DISTINCT [CorrelationID]
        FROM
            [dbo].[ActivityLog]
        WHERE 
            ComponentID=@JobId AND
            CorrelationId IS NOT NULL AND
            ItemUniqueId LIKE @ItemId)
ORDER BY
    [js].[Start] DESC";
        private const string sql_GetTopJobStatistics = @"
SELECT TOP (@ItemCount) 
    [js].*,
    (SELECT TOP (1) [Activity] FROM [dbo].[ActivityLog] WHERE [CorrelationID]=js.CorrelationId AND ([Activity]='JobCompleted' OR [Activity]='JobFailed')) AS JobState
FROM
    [dbo].[JobStatistics] AS [js]
INNER JOIN 
    [dbo].[Components] AS [c] ON ([js].[JobName]=[c].[InternalName])
WHERE 
    [c].[ComponentId]=@JobId";

        private const string sql_GetJobTimeoutState = @"
SELECT TOP(1)
    IIF ([Timestamp] > DATEADD(minute, 0-@TimeoutThreshold, GETUTCDATE()), 0, 1) AS [TimeoutThresholdReached]
  FROM [dbo].[ActivityLog]
  WHERE [CorrelationID] = @CorrelationId
  ORDER BY [Timestamp]";

        private const string sql_GetWeeklyMCStatistics = @"
SELECT DATEADD(dd,  0, DATEADD(ww, DATEDIFF(ww, 0, DATEADD(dd, -1, [LastUpdatedTime])), 0)) AS [Week],
      COUNT(ID) AS [Items]
  FROM [dbo].[Notifications]
  WHERE
    [Type]=@messageType AND
    [LastUpdatedTime] >= DATEADD(dd,  0, DATEADD(ww, DATEDIFF(ww, 0, DATEADD(dd, -1, GETUTCDATE())) - 12, 0)) AND
    [LastUpdatedTime] <= DATEADD(dd,  7, DATEADD(ww, DATEDIFF(ww, 0, DATEADD(dd, -1, GETUTCDATE())) - 1, 0))
  GROUP BY DATEADD(dd,  0, DATEADD(ww, DATEDIFF(ww, 0, DATEADD(dd, -1, [LastUpdatedTime])), 0))
  ORDER BY DATEADD(dd,  0, DATEADD(ww, DATEDIFF(ww, 0, DATEADD(dd, -1, [LastUpdatedTime])), 0))
";
        private const string sql_GetMCStatisticsForPast7Days = @"
SELECT COUNT(ID) AS [Items],
       [svc].[Service] AS [Service]
FROM [dbo].[Notifications]
CROSS APPLY
    OPENJSON ([Data], '$.services') WITH([Service] nvarchar(256) '$') [svc]
WHERE
    [Type]='SERVICEUPDATEMESSAGE' AND
    [LastUpdatedTime] >= DATEADD(dd,  -7, GETUTCDATE()) AND
    [LastUpdatedTime] <= GETUTCDATE()
GROUP BY
    [svc].[Service]
ORDER BY
    [svc].[Service]
";
        private const string sql_GetAzureServiceHealthAlerts = "SELECT * FROM [dbo].[vw_GetCommunications] WHERE [Type]='AZURESERVICEHEALTHALERT'";
        private const string sql_GetAzureUpdates = "SELECT * FROM [dbo].[vw_GetCommunications] WHERE [Type]='AZUREUPDATE'";
        private const string sql_GetAzureUpdatesActive = @"
SELECT
    *
FROM [dbo].[vw_GetCommunications]
WHERE
    [Type] = 'AZUREUPDATE'
    AND json_value([Data],'$.releaseStatus') <> 'Launched'
UNION ALL
SELECT
    *
FROM
    [dbo].[vw_GetCommunications]
WHERE
    [Type] = 'AZUREUPDATE'
    AND LastUpdatedTime > DATEADD(year, -2, GETUTCDATE())
    AND (json_value([Data],'$.releaseStatus') IS NULL)
ORDER BY LastUpdatedTime DESC;
";
        private const string sql_GetD365PowerPlatformReleases = "SELECT * FROM [dbo].[AllD365PPCommsDetailed]";
        private const string sql_GetMonthlyActiveUsersReport = "SELECT * FROM [dbo].[vw_GetMonthlyUsageReport]";
        private const string sql_GetLastNotificationSummary = "SELECT TOP(1) * FROM [dbo].[SummaryCache] WHERE MessageId=@MessageId ORDER BY [Timestamp] DESC";
        private const string sql_GetLastNotificationSummaryList = "SELECT * FROM [dbo].[vw_SummaryCacheNewest]";
        private const string sql_GetDatabaseVersion = "SELECT TOP(1) [Version] FROM [dbo].[Versions] ORDER BY [Version] DESC";
        private const string sql_GetViewpoint = "SELECT * FROM [dbo].[fn_GetViewpointCollection](@CommunicationIds)";
        private const string sql_sp_SetConfigurationValue = "[dbo].[proc_SetConfigValue]";
        private const string sql_sp_CacheTranslation = "[dbo].[proc_CacheTranslation]";
        private const string sql_sp_PublishMessage = "[dbo].[proc_PublishNotification]";
        private const string sql_sp_UnpublishMessage = "[dbo].[proc_UnpublishNotification]";
        private const string sql_sp_UpdateSyncConfigEntry = "[dbo].[proc_AddSyncConfig]";
        private const string sql_sp_UpdateConnector = "[dbo].[proc_AddConnector]";
        private const string sql_sp_DeleteConnector = "[dbo].[proc_DeleteConnector]";
        private const string sql_sp_UpdateRoute = "[dbo].[proc_AddRoute]";
        private const string sql_sp_DeleteRoute = "[dbo].[proc_DeleteRoute]";
        private const string sql_sp_AddCustomAction = "[dbo].[proc_AddCustomAction]";
        private const string sql_sp_DeleteCustomAction = "[dbo].[proc_DeleteCustomAction]";
        private const string sql_sp_AddActivityLogRecord = "[dbo].[proc_AddActivityLogRecord]";
        private const string sql_sp_SetArchiveFlag = "[dbo].[proc_SetArchiveFlag]";
        private const string sql_sp_SetUserProfileProperties = "[dbo].[proc_SetUserProfileProperties]";
        private const string sql_sp_SetViewpointReadFlag = "[dbo].[proc_SetViewpointReadFlag]";
        private const string sql_sp_SetViewpointArchiveFlag = "[dbo].[proc_SetViewpointArchiveFlag]";
        private const string sql_sp_SetViewpointFavoriteFlag = "[dbo].[proc_SetViewpointFavoriteFlag]";
        private const string sql_sp_AddTagDefinition = "[dbo].[proc_AddTagDefinition]";
        private const string sql_sp_RemoveTagDefinition = "[dbo].[proc_RemoveTagDefinition]";
        private const string sql_sp_MoveTagDefinition = "[dbo].[proc_MoveTagDefinition]";
        private const string sql_sp_AddTag = "[dbo].[proc_AddTag]";
        private const string sql_sp_RemoveTag = "[dbo].[proc_RemoveTag]";

        private MSSQLServerConnector m_sqlConn;
        public static string GetSqlAzureConnectionString(string name)
        {
            string conStr = System.Environment.GetEnvironmentVariable($"ConnectionStrings:{name}", EnvironmentVariableTarget.Process);
            if (string.IsNullOrEmpty(conStr)) // Azure Functions App Service naming convention
                conStr = System.Environment.GetEnvironmentVariable($"SQLAZURECONNSTR_{name}", EnvironmentVariableTarget.Process);
            return conStr;
        }

        public MSSHNotificationDatabase ()
        {
            string connectionString = GlobalConfiguration.Instance.NotificationDatabaseConnectionString;
            m_sqlConn = new MSSQLServerConnector(connectionString);
        }

        public MSSHNotificationDatabase(string connectionString)
        {

            m_sqlConn = new MSSQLServerConnector(connectionString);
        }

        public string GetSchemaVersion()
        {
            List<MSPropertyBag> result = new List<MSPropertyBag>();
            string version = string.Empty;

            MSPropertyBag parameters = new ();
            result = m_sqlConn.GetDataSet(sql_GetDatabaseVersion, parameters);
            
            if (result != null && result.Count > 0)
            {
                MSPropertyBag versionData = result[0];
                version = versionData["Version"] is DBNull ? string.Empty : (string)versionData["Version"];
            }

            return version;
        }
        public MSServiceNotification GetServiceNotification(string id)
        {
            MSServiceNotification notification = new MSServiceNotification();
            GetNotification(id, notification);
            return notification;
        }

        public object GetConfigurationValue(string name)
        {
            MSPropertyBag parameters = new()
            {
                { "Name", name }
            };
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetConfigurationValue, parameters);

            if (result != null && result.Count > 0)
            {
                // we shold get max. one record anyway, fetching the first record from the result set
                MSPropertyBag configData = result[0];
                object valueObj = null;
                if (configData["SerializedValue"] is not DBNull)
                {
                    string value = configData["SerializedValue"] is DBNull ? string.Empty : (string)configData["SerializedValue"];
                    try
                    {
                        valueObj = JsonConvert.DeserializeObject<object>(value);
                    } 
                    catch
                    {
                        valueObj = value;
                    }
                }
                return valueObj;
            }
            else
                return null;
        }

        public T GetConfigurationValue<T>(string name)
        {
            MSPropertyBag parameters = new()
            {
                { "Name", name }
            };
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetConfigurationValue, parameters);

            if (result != null && result.Count > 0)
            {
                // we shold get max. one record anyway, fetching the first record from the result set
                MSPropertyBag configData = result[0];
                T valueObj = default;
                if (configData["SerializedValue"] is not DBNull)
                {
                    string value = configData["SerializedValue"] is DBNull ? string.Empty : (string)configData["SerializedValue"];
                    try
                    {
                        valueObj = JsonConvert.DeserializeObject<T>(value);
                    }
                    catch
                    {
                        valueObj = default;
                    }
                }
                return valueObj;
            }
            else
                return default;
        }

        public void SetConfigurationValue(string name, object value)
        {
            MSPropertyBag parameters = new()
            {
                { "Key", name },
                { "Value", value == null ? null : JsonConvert.SerializeObject(value) },
                { "DataType", value == null ? "" : value.GetType().FullName }
            };
            m_sqlConn.ExecuteStoredProcedure(sql_sp_SetConfigurationValue, parameters);
        }

        public void GetNotification(string id, MSServiceNotification notification)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ID", id }
            };
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetNotification, parameters);
            if (result != null && result.Count > 0)
            {
                // we shold get max. one record anyway, fetching the first record from the result set
                MSPropertyBag instanceData = result[0];
                notification.m_Id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"];
                notification.m_LastUpdatedTime = instanceData["LastUpdatedTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["LastUpdatedTime"];
                notification.m_Data = instanceData["Data"] is DBNull ? string.Empty : (string)instanceData["Data"];
                notification.m_public = instanceData["Public"] is DBNull ? false : (bool)instanceData["Public"];
                notification.m_comments = instanceData["Comments"] is DBNull ? string.Empty : (string)instanceData["Comments"];
                notification.m_WorkItemID = instanceData["WorkItemID"] is DBNull ? string.Empty : (string)instanceData["WorkItemID"];
                notification.m_WorkItemURL = instanceData["WorkItemURL"] is DBNull ? string.Empty : (string)instanceData["WorkItemURL"];
                notification.m_ExtendedProperties = instanceData["ExtendedProperties"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ExtendedProperties"]);
                notification.m_State = instanceData["State"] is DBNull ? string.Empty : (string)instanceData["State"];
            }
        }

        public void GetNotification(string id, string type, MSServiceNotification notification)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ID", id },
                { "Type", type }
            };
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetNotification + " AND [Type]=@Type", parameters);
            if (result != null && result.Count > 0)
            {
                // we shold get max. one record anyway, fetching the first record from the result set
                MSPropertyBag instanceData = result[0];
                notification.m_Id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"];
                notification.m_LastUpdatedTime = instanceData["LastUpdatedTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["LastUpdatedTime"];
                notification.m_Data = instanceData["Data"] is DBNull ? string.Empty : (string)instanceData["Data"];
                notification.m_public = instanceData["Public"] is DBNull ? false : (bool)instanceData["Public"];
                notification.m_comments = instanceData["Comments"] is DBNull ? string.Empty : (string)instanceData["Comments"];
                notification.m_WorkItemID = instanceData["WorkItemID"] is DBNull ? string.Empty : (string)instanceData["WorkItemID"];
                notification.m_WorkItemURL = instanceData["WorkItemURL"] is DBNull ? string.Empty : (string)instanceData["WorkItemURL"];
                notification.m_ExtendedProperties = instanceData["ExtendedProperties"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ExtendedProperties"]);
                notification.m_State = instanceData["State"] is DBNull ? string.Empty : (string)instanceData["State"];
            }
        }

        public void GetNotifications(List<string> ids, List<MSServiceNotification> notifications)
        {
            MSPropertyBag parameters = new MSPropertyBag();

            string sqlQuery = sql_GetNotifications;

            if (ids != null && ids.Count > 0)
            {
                string idList = "'" + string.Join("','", ids) + "'";
                sqlQuery += "WHERE [ID] IN (" + idList + ")";
            }
            
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sqlQuery, parameters);

            if (result != null && result.Count > 0)
            {               
                foreach (MSPropertyBag instanceData in result)
                {
                    MSServiceNotification notification = new MSServiceNotification();
                    notification.m_Id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"];
                    notification.m_LastUpdatedTime = instanceData["LastUpdatedTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["LastUpdatedTime"];
                    notification.m_Data = instanceData["Data"] is DBNull ? string.Empty : (string)instanceData["Data"];
                    notification.m_public = instanceData["Public"] is DBNull ? false : (bool)instanceData["Public"];
                    notification.m_comments = instanceData["Comments"] is DBNull ? string.Empty : (string)instanceData["Comments"];
                    notification.m_WorkItemID = instanceData["WorkItemID"] is DBNull ? string.Empty : (string)instanceData["WorkItemID"];
                    notification.m_WorkItemURL = instanceData["WorkItemURL"] is DBNull ? string.Empty : (string)instanceData["WorkItemURL"];
                    notification.m_ExtendedProperties = instanceData["ExtendedProperties"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ExtendedProperties"]);
                    notification.m_State = instanceData["State"] is DBNull ? string.Empty : (string)instanceData["State"];
                    notifications.Add(notification);
                }
            }
        }

        public void GetIncidents(string workloads, string startDate, string endDate, List<MSServiceIncident> incidents)
        {
            string[] workloadArray = string.IsNullOrWhiteSpace(workloads) ? new string[0] :  workloads.Split(',', StringSplitOptions.RemoveEmptyEntries);
            string query = sql_IncidentTimeline;
            if (workloadArray.Length > 0)
            {
                query += " WHERE ";
                foreach (string w in workloadArray)
                {
                    query += $"[Workload] LIKE '%{w.Trim()}%' OR";
                }
                query = query.Substring(0, query.Length - 3);

            }

            if (!string.IsNullOrWhiteSpace(startDate))
            {
                DateTime dt = DateTime.UtcNow;
                if (DateTime.TryParse(startDate, out dt))
                {
                    if (query.ToUpper().Contains("WHERE"))
                        query += $" AND [StartTime] >= '{startDate}'";
                    else
                        query += $" WHERE [StartTime] >= '{startDate}'";
                }
            }

            if (!string.IsNullOrWhiteSpace(endDate))
            {
                DateTime dt = DateTime.UtcNow;
                if (DateTime.TryParse(endDate, out dt))
                {
                    if (query.ToUpper().Contains("WHERE"))
                        query += $" AND [StartTime] <= '{endDate}'";
                    else
                        query += $" WHERE [StartTime] >= '{endDate}'";
                }
            }

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query);
            if (result != null && result.Count > 0)
            {
                foreach (var r in result)
                {
                    MSPropertyBag instanceData = r;
                    MSServiceIncident incident = new MSServiceIncident();

                    incident.m_Id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"];
                    incident.m_Title = instanceData["Title"] is DBNull ? string.Empty : (string)instanceData["Title"];
                    incident.m_Workload = instanceData["Workload"] is DBNull ? string.Empty : (string)instanceData["Workload"];
                    incident.m_StartTime = instanceData["StartTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["StartTime"];
                    incident.m_EndTime = instanceData["EndTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["EndTime"];
                    incidents.Add(incident);
                }               
            }
        }

        public void GetServiceWorkloads(List<string> workloads)
        {
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetServiceWorkloads);
            if (result != null && result.Count > 0)
            {
                foreach (var r in result)
                {
                    string workload = r["Workload"] is DBNull ? string.Empty : (string)r["Workload"];
                    if (!string.IsNullOrWhiteSpace(workload))
                        workloads.Add(workload);
                }
            }
        }

        public void GetRoadmapNotificationCollection(MSRoadmapNotificationCollection notifications)
        {
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetFutureRoadmapItems);
            if (result != null && result.Count > 0)
            {
                foreach (var r in result)
                {
                    MSPropertyBag instanceData = r;
                    MSRoadmapNotification notification = new MSRoadmapNotification
                    {
                        m_Id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"],
                        m_Title = instanceData["Title"] is DBNull ? string.Empty : (string)instanceData["Title"],
                        m_Status = instanceData["Status"] is DBNull ? string.Empty : (string)instanceData["Status"],
                        m_Link = instanceData["Link"] is DBNull ? string.Empty : (string)instanceData["Link"],
                        m_MoreInfoLink = instanceData["MoreInfoLink"] is DBNull ? string.Empty : (string)instanceData["MoreInfoLink"],
                        m_Description = instanceData["Description"] is DBNull ? string.Empty : (string)instanceData["Description"],
                        m_AvailabilityDate = instanceData["AvailabilityDate"] is DBNull ? string.Empty : (string)instanceData["AvailabilityDate"],
                        m_AvailabilityFrom = instanceData["AvailabilityFrom"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["AvailabilityFrom"],
                        m_AvailabilityTo = instanceData["AvailabilityTo"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["AvailabilityTo"],
                        m_PublicPreviewDate = instanceData["PublicPreviewDate"] is DBNull ? string.Empty : (string)instanceData["PublicPreviewDate"],
                        m_PublicPreviewFrom = instanceData["PublicPreviewFrom"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["PublicPreviewFrom"],
                        m_PublicPreviewTo = instanceData["PublicPreviewTo"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["PublicPreviewTo"],
                        m_Published = instanceData["Published"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["Published"],
                        m_LastUpdated = instanceData["LastUpdated"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["LastUpdated"],
                        m_Category = instanceData["Category"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Category"]),
                        m_Products = instanceData["Products"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Products"]),
                        m_Tags = instanceData["Tags"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Tags"]),
                        m_CloudInstances = instanceData["CloudInstances"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["CloudInstances"]),
                        m_ReleasePhase = instanceData["ReleasePhase"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ReleasePhase"]),
                        m_Platforms = instanceData["Platforms"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Platforms"]),
                        m_WorkItemID = instanceData["WorkItemID"] is DBNull ? string.Empty : (string)instanceData["WorkItemID"],
                        m_WorkItemURL = instanceData["WorkItemURL"] is DBNull ? string.Empty : (string)instanceData["WorkItemURL"],
                        m_ExtendedProperties = instanceData["ExtendedProperties"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ExtendedProperties"])
                };

                    notifications.Add(notification);
                }
            }
        }

        public MSRoadmapNotification GetRoadmapNotification(string id)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ID", id }
            };

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetRoadmapItem, parameters);

            if (result != null && result.Count > 0)
            {
                MSPropertyBag instanceData = result[0];
                MSRoadmapNotification notification = new MSRoadmapNotification
                {
                    m_Id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"],
                    m_Title = instanceData["Title"] is DBNull ? string.Empty : (string)instanceData["Title"],
                    m_Status = instanceData["Status"] is DBNull ? string.Empty : (string)instanceData["Status"],
                    m_Link = instanceData["Link"] is DBNull ? string.Empty : (string)instanceData["Link"],
                    m_MoreInfoLink = instanceData["MoreInfoLink"] is DBNull ? string.Empty : (string)instanceData["MoreInfoLink"],
                    m_Description = instanceData["Description"] is DBNull ? string.Empty : (string)instanceData["Description"],
                    m_AvailabilityDate = instanceData["AvailabilityDate"] is DBNull ? string.Empty : (string)instanceData["AvailabilityDate"],
                    m_AvailabilityFrom = instanceData["AvailabilityFrom"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["AvailabilityFrom"],
                    m_AvailabilityTo = instanceData["AvailabilityTo"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["AvailabilityTo"],
                    m_PublicPreviewDate = instanceData["PublicPreviewDate"] is DBNull ? string.Empty : (string)instanceData["PublicPreviewDate"],
                    m_PublicPreviewFrom = instanceData["PublicPreviewFrom"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["PublicPreviewFrom"],
                    m_PublicPreviewTo = instanceData["PublicPreviewTo"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["PublicPreviewTo"],
                    m_Published = instanceData["Published"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["Published"],
                    m_LastUpdated = instanceData["LastUpdated"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["LastUpdated"],
                    m_Category = instanceData["Category"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Category"]),
                    m_Products = instanceData["Products"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Products"]),
                    m_Tags = instanceData["Tags"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Tags"]),
                    m_CloudInstances = instanceData["CloudInstances"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["CloudInstances"]),
                    m_ReleasePhase = instanceData["ReleasePhase"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ReleasePhase"]),
                    m_Platforms = instanceData["Platforms"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Platforms"]),
                    m_WorkItemID = instanceData["WorkItemID"] is DBNull ? string.Empty : (string)instanceData["WorkItemID"],
                    m_WorkItemURL = instanceData["WorkItemURL"] is DBNull ? string.Empty : (string)instanceData["WorkItemURL"],
                    m_ExtendedProperties = instanceData["ExtendedProperties"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ExtendedProperties"])
                };

                return notification;
            }
            else
                return null;
        }

        public MSPropertyBag GetRoadmapStatistics()
        {
            MSPropertyBag stats = new MSPropertyBag();

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetRoadmapStatistics);
            if (result != null && result.Count > 0)
            {
                MSPropertyBag statsData = result[0];
                stats["inDevelopment"] = statsData["inDevelopment"] is DBNull ? 0 : (int)statsData["inDevelopment"];
                stats["rollingOut"] = statsData["rollingOut"] is DBNull ? 0 : (int)statsData["rollingOut"];
            }
            else
            {
                stats["inDevelopment"] = 0;
                stats["rollingOut"] = 0;
            }

            return stats;
        }

        public void GetPublicIncidentCollection(MSPublicIncidentCollection incidents)
        {
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetPublicIncidents);
            if (result != null && result.Count > 0)
            {
                foreach (var r in result)
                {
                    MSPropertyBag instanceData = r;
                    MSPublicIncident incident = new MSPublicIncident
                    {
                        _id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"],
                        _title = instanceData["Title"] is DBNull ? string.Empty : (string)instanceData["Title"],
                        _service = instanceData["Service"] is DBNull ? string.Empty : (string)instanceData["Service"],
                        _startTime = instanceData["StartTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["StartTime"],
                        _lastModified = instanceData["lastModifiedDateTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["lastModifiedDateTime"],
                        _comments = instanceData["Comments"] is DBNull ? string.Empty : (string)instanceData["Comments"]
                    };
                    if (string.IsNullOrWhiteSpace(incident._comments))
                    {
                        incident._comments = "No additional information provided. Please wait until the incident is resolved.";
                    }
                    incidents.Add(incident);
                }
            }
        }

        private List<string>DeserializeServiceList(string serviceListSerialized)
        {
            List<string> serviceList = new List<string>();
            if (!string.IsNullOrWhiteSpace(serviceListSerialized))
            {
                serviceListSerialized = "{ \"services\": " + serviceListSerialized + " }";
                MSPublicMessageServicesList svcList = JsonConvert.DeserializeObject<MSPublicMessageServicesList>(serviceListSerialized);
                if (svcList?.services?.Count > 0)
                {
                    serviceList = svcList.services;
                }
            }

            return serviceList;
        }

        public void GetPublicMessageCollection(MSPublicMessageCollection messages)
        {
            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetPublicMessages);
            if (result != null && result.Count > 0)
            {
                foreach (var r in result)
                {
                    MSPropertyBag instanceData = r;
                    string serviceListSerialized = instanceData["Services"] is DBNull ? string.Empty : (string)instanceData["Services"];
                    List<string> serviceList = DeserializeServiceList(serviceListSerialized);

                    MSPublicMessage message = new MSPublicMessage
                    {
                        _id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"],
                        _title = instanceData["Title"] is DBNull ? string.Empty : (string)instanceData["Title"],
                        _services = serviceList,
                        _startTime = instanceData["StartTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["StartTime"],
                        _lastModified = instanceData["lastModifiedDateTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["lastModifiedDateTime"],
                        _comments = instanceData["Comments"] is DBNull ? string.Empty : (string)instanceData["Comments"]
                    };
                    
                    messages.Add(message);
                }
            }
        }

        public void GetPublicMessage(string id, MSPublicMessage message)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ID", id }
            };

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetPublicMessage, parameters);
            if (result != null && result.Count > 0)
            {
                var r = result[0];
                
                MSPropertyBag instanceData = r;
                string serviceListSerialized = instanceData["Services"] is DBNull ? string.Empty : (string)instanceData["Services"];
                List<string> serviceList = DeserializeServiceList(serviceListSerialized);

                message._id = instanceData["ID"] is DBNull ? string.Empty : (string)instanceData["ID"];
                message._title = instanceData["Title"] is DBNull ? string.Empty : (string)instanceData["Title"];
                message._services = serviceList;
                message._startTime = instanceData["StartTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["StartTime"];
                message._lastModified = instanceData["lastModifiedDateTime"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["lastModifiedDateTime"];
                message._comments = instanceData["Comments"] is DBNull ? string.Empty : (string)instanceData["Comments"];
                message._content = instanceData["Content"] is DBNull ? string.Empty : (string)instanceData["Content"];
            }
        }

        public void PublishMessage(string id, string type, string comments, MSUserInfo userInfo)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "Id", id },
                { "Type", type },
                { "Comments", comments },
                { "LastModifiedBy", JsonConvert.SerializeObject(userInfo) }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_PublishMessage, parameters);
        }

        public void UnpublishMessage(string id, string type, MSUserInfo userInfo)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "Id", id },
                { "Type", type },
                { "LastModifiedBy", JsonConvert.SerializeObject(userInfo) }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_UnpublishMessage, parameters);
        }

        private string ComputeMD5Hash(string Message)
        {
            var md5 = MD5.Create();
            var utf8 = new UTF8Encoding();
            string hash = BitConverter.ToString(md5.ComputeHash(utf8.GetBytes(Message)));
            hash = hash.Replace("-", "");

            return hash;
        }

        public MSTranslation GetTranslationFromCache(string Message, string Language)
        {
            MSTranslation result = null;

            string hash = ComputeMD5Hash(Message);

            MSPropertyBag parameters = new MSPropertyBag
            {
                { "Hash", hash },
                { "Language", Language }
            };

            List<MSPropertyBag> queryResults = m_sqlConn.GetDataSet(sql_GetTranslationFromCache, parameters);

            if (queryResults != null && queryResults.Count > 0)
            {
                // we shold get max. one record anyway, fetching the first record from the result set
                MSPropertyBag translation = queryResults[0];
                result = new MSTranslation
                {
                    _hash = translation["Hash"] is DBNull ? string.Empty : (string)translation["Hash"],
                    _language = translation["Language"] is DBNull ? string.Empty : (string)translation["Language"],
                    _message = translation["Message"] is DBNull ? string.Empty : (string)translation["Message"]
                };

            }

            return result;
        }

        public List<MSTranslationCollection> GetTranslationFromCache(string[] Messages, string Language)
        {
            List<MSTranslationCollection> result = new List<MSTranslationCollection>();
            Dictionary<string, string> hashMapping = new Dictionary<string, string>();
            List<string> hashlist = new List<string>();
            string hashParam = string.Empty;

            foreach (string message in Messages)
            {
                string hash = ComputeMD5Hash(message);
                hashlist.Add(hash);
                hashMapping.Add(hash, message);
            }

            hashParam = string.Join("', '", hashlist);
            if (!string.IsNullOrWhiteSpace(hashParam))
            {
                hashParam = "'" + hashParam + "'";
            }           

            MSPropertyBag parameters = new MSPropertyBag
            {
                { "Language", Language }
            };

            List<MSPropertyBag> queryResults = m_sqlConn.GetDataSet(string.Format(sql_GetTranslationsFromCache, hashParam), parameters);

            if (queryResults != null && queryResults.Count > 0)
            {
                foreach (MSPropertyBag translation in queryResults)
                {
                    MSTranslation cacheItem = new MSTranslation
                    {
                        _hash = translation["Hash"] is DBNull ? string.Empty : (string)translation["Hash"],
                        _language = translation["Language"] is DBNull ? string.Empty : (string)translation["Language"],
                        _message = translation["Message"] is DBNull ? string.Empty : (string)translation["Message"]
                    };

                    MSTranslationCollection translationCollection = new MSTranslationCollection();
                    translationCollection.OriginalMessage = hashMapping[cacheItem.Hash];
                    translationCollection.Translations.Add(cacheItem.Language, cacheItem);
                    result.Add(translationCollection);
                }

            }

            return result;
        }

        public void CacheTranslation(string originalMessage, MSTranslation translation)
        {
            translation.Hash = ComputeMD5Hash(originalMessage);

            MSPropertyBag parameters = new MSPropertyBag
            {
                { "Hash", translation.Hash },
                { "Language", translation.Language },
                { "Message", translation.Message },
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_CacheTranslation, parameters);
        }

        public void GetSyncConfigEntry(string component, string element, MSSyncConfigEntry config)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "Component", component },
                { "Element", element }
            };

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetSyncConfigEntry, parameters);
            if (result != null && result.Count > 0)
            {
                var r = result[0];

                MSPropertyBag instanceData = r;
                
                config._id = instanceData["Id"] is DBNull ? -1 : (Int32)instanceData["Id"];
                config._component = instanceData["Component"] is DBNull ? string.Empty : (string)instanceData["Component"];
                config._element = instanceData["Element"] is DBNull ? string.Empty : (string)instanceData["Element"];
                config._config = instanceData["Config"] is DBNull ? string.Empty : (string)instanceData["Config"];
                config._created = instanceData["Created"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["Created"];
                config._createdBy = instanceData["CreatedBy"] is DBNull ? string.Empty : (string)instanceData["CreatedBy"];
                config._modified = instanceData["Modified"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["Modified"];
                config._modifiedBy = instanceData["ModifiedBy"] is DBNull ? string.Empty : (string)instanceData["ModifiedBy"];
            } else
            {
                config._component = component;
                config._element = element;
            }
        }

        public void UpdateSyncConfigEntry(MSSyncConfigEntry configEntry, MSUserInfo userInfo)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "Component", configEntry.Component },
                { "Element", configEntry.Element},
                { "Config", configEntry.Config },
                { "User", userInfo.UserName }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_UpdateSyncConfigEntry, parameters);
        }

        public List<MSLicenseStatistics> GetLicenseStatistics(string sku = "")
        {
            List<MSLicenseStatistics> statsCollection = new List<MSLicenseStatistics>();
            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetLastLicenseStatistics;
            if (!string.IsNullOrWhiteSpace(sku))
            {
                query += " WHERE [SkuPartNumber]=@Sku";
                parameters.Add("Sku", sku);
            }

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);

            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSLicenseStatistics stats = new MSLicenseStatistics
                    {
                        DisplayName = instanceData["DisplayName"] is DBNull ? string.Empty : (string)instanceData["DisplayName"],
                        SkuPartNumber = instanceData["SkuPartNumber"] is DBNull ? string.Empty : (string)instanceData["SkuPartNumber"],
                        Enabled = instanceData["Enabled"] is DBNull ? 0 : (int)instanceData["Enabled"],
                        ConsumedUnits = instanceData["ConsumedUnits"] is DBNull ? 0 : (int)instanceData["ConsumedUnits"],
                        AvailableUnits = instanceData["AvailableUnits"] is DBNull ? 0 : (int)instanceData["AvailableUnits"],
                        Warning = instanceData["Warning"] is DBNull ? 0 : (int)instanceData["Warning"],
                        Suspended = instanceData["Suspended"] is DBNull ? 0 : (int)instanceData["Suspended"],
                        Date = instanceData["Date"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["Date"],
                    };
                    statsCollection.Add(stats);
                }
            }

            return statsCollection;
        }

        public List<MSLicenseStatisticsExt> GetAllLicenseStatistics(string sku = "")
        {
            List<MSLicenseStatisticsExt> statsCollection = new List<MSLicenseStatisticsExt>();
            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetAllLicenseStatistics;
            if (!string.IsNullOrWhiteSpace(sku))
            {
                query += " AND [SkuPartNumber]=@Sku";
                parameters.Add("Sku", sku);
            }

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSLicenseStatisticsExt stats = new MSLicenseStatisticsExt
                    {
                        DisplayName = instanceData["DisplayName"] is DBNull ? string.Empty : (string)instanceData["DisplayName"],
                        SkuPartNumber = instanceData["SkuPartNumber"] is DBNull ? string.Empty : (string)instanceData["SkuPartNumber"],
                        Enabled = instanceData["Enabled"] is DBNull ? 0 : (int)instanceData["Enabled"],
                        ConsumedUnits = instanceData["ConsumedUnits"] is DBNull ? 0 : (int)instanceData["ConsumedUnits"],
                        Suspended = instanceData["Suspended"] is DBNull ? 0 : (int)instanceData["Suspended"],
                        Warning = instanceData["Warning"] is DBNull ? 0 : (int)instanceData["Warning"],
                        Timestamp = instanceData["Timestamp"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["Timestamp"],
                    };
                    statsCollection.Add(stats);
                }
            }

            return statsCollection;
        }

        public List<MSLicenseForecast> GetLicenseForecast(string sku = "")
        {
            List<MSLicenseForecast> statsCollection = new List<MSLicenseForecast>();
            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetAllLicensesForecast;
            if (!string.IsNullOrWhiteSpace(sku))
            {
                query += " WHERE [SkuPartNumber]=@Sku";
                parameters.Add("Sku", sku);
            }

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSLicenseForecast stats = new MSLicenseForecast
                    {
                        SkuPartNumber = instanceData["SkuPartNumber"] is DBNull ? string.Empty : (string)instanceData["SkuPartNumber"],
                        NewHires = instanceData["NewHires"] is DBNull ? 0 : (int)instanceData["NewHires"],
                        Leavers = instanceData["Leavers"] is DBNull ? 0 : (int)instanceData["Leavers"],
                        Balance = instanceData["Balance"] is DBNull ? 0 : (int)instanceData["Balance"],
                        Month = instanceData["Month"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["Month"],
                    };

                    statsCollection.Add(stats);
                }
            }

            return statsCollection;
        }

        public List<MSConnector> GetConnectors(string connectorType = "")
        {
            List<MSConnector> connectorCollection = new List<MSConnector>();
            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetConnectors;
            if (!string.IsNullOrWhiteSpace(connectorType))
            {
                query += " WHERE [ConnectorType]=@ConnectorType";
                parameters.Add("ConnectorType", connectorType);
            }

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSConnector connector = new MSConnector
                    {
                        ConnectorId = instanceData["ConnectorId"] is DBNull ? Guid.Empty : (Guid)instanceData["ConnectorId"],
                        Name = instanceData["Name"] is DBNull ? string.Empty : (string)instanceData["Name"],
                        Type = instanceData["Type"] is DBNull ? Guid.Empty : (Guid)instanceData["Type"],
                        ConnectorTypeName = instanceData["ConnectorTypeName"] is DBNull ? string.Empty : (string)instanceData["ConnectorTypeName"],
                        ConnectorType = instanceData["ConnectorType"] is DBNull ? string.Empty : (string)instanceData["ConnectorType"],
                        Icon = instanceData["Icon"] is DBNull ? string.Empty : (string)instanceData["Icon"],
                        System = instanceData["System"] is DBNull ? false : (bool)instanceData["System"],
                        Hidden = instanceData["Hidden"] is DBNull ? false : (bool)instanceData["Hidden"],
                        Configuration = instanceData["Configuration"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Configuration"]),
                        ParameterDefinition = instanceData["ParameterDefinition"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ParameterDefinition"]),
                    };
                    connectorCollection.Add(connector);
                }
            }

            return connectorCollection;
        }

        public List<MSConnector> GetConnectors(Guid definitionId)
        {
            List<MSConnector> connectorCollection = new List<MSConnector>();
            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetConnectors;
            query += " WHERE [Type]=@Type";
            parameters.Add("Type", definitionId);

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSConnector connector = new MSConnector
                    {
                        ConnectorId = instanceData["ConnectorId"] is DBNull ? Guid.Empty : (Guid)instanceData["ConnectorId"],
                        Name = instanceData["Name"] is DBNull ? string.Empty : (string)instanceData["Name"],
                        Type = instanceData["Type"] is DBNull ? Guid.Empty : (Guid)instanceData["Type"],
                        ConnectorTypeName = instanceData["ConnectorTypeName"] is DBNull ? string.Empty : (string)instanceData["ConnectorTypeName"],
                        ConnectorType = instanceData["ConnectorType"] is DBNull ? string.Empty : (string)instanceData["ConnectorType"],
                        Icon = instanceData["Icon"] is DBNull ? string.Empty : (string)instanceData["Icon"],
                        System = instanceData["System"] is DBNull ? false : (bool)instanceData["System"],
                        Hidden = instanceData["Hidden"] is DBNull ? false : (bool)instanceData["Hidden"],
                        Configuration = instanceData["Configuration"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Configuration"]),
                        ParameterDefinition = instanceData["ParameterDefinition"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ParameterDefinition"]),
                    };
                    connectorCollection.Add(connector);
                }
            }

            return connectorCollection;
        }

        public MSConnector GetConnector(Guid connectorId)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            MSConnector connector = null;

            string query = sql_GetConnectors;
            
            query += " WHERE [ConnectorId]=@ConnectorId";
            parameters.Add("ConnectorId", connectorId);

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                MSPropertyBag instanceData = result[0];
               
                connector = new MSConnector
                {
                    ConnectorId = instanceData["ConnectorId"] is DBNull ? Guid.Empty : (Guid)instanceData["ConnectorId"],
                    Name = instanceData["Name"] is DBNull ? string.Empty : (string)instanceData["Name"],
                    Type = instanceData["Type"] is DBNull ? Guid.Empty : (Guid)instanceData["Type"],
                    ConnectorTypeName = instanceData["ConnectorTypeName"] is DBNull ? string.Empty : (string)instanceData["ConnectorTypeName"],
                    ConnectorType = instanceData["ConnectorType"] is DBNull ? string.Empty : (string)instanceData["ConnectorType"],
                    Icon = instanceData["Icon"] is DBNull ? string.Empty : (string)instanceData["Icon"],
                    System = instanceData["System"] is DBNull ? false : (bool)instanceData["System"],
                    Hidden = instanceData["Hidden"] is DBNull ? false : (bool)instanceData["Hidden"],
                    Configuration = instanceData["Configuration"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Configuration"]),
                    ParameterDefinition = instanceData["ParameterDefinition"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ParameterDefinition"]),
                };
            }

            return connector;
        }

        public void UpdateConnector(MSConnector connector)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ConnectorId", connector.ConnectorId },
                { "Name", connector.Name },
                { "Type", connector.Type },
                { "Configuration", connector.Configuration.ToString() }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_UpdateConnector, parameters);
        }

        public void DeleteConnector(MSConnector connector)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ConnectorId", connector.ConnectorId }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_DeleteConnector, parameters);
        }

        public List<MSConnectorDefinition> GetConnectorDefinitions(string connectorType = "")
        {
            List<MSConnectorDefinition> connectorDefCollection = new List<MSConnectorDefinition>();
            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetConnectorDefinitions;
            if (!string.IsNullOrWhiteSpace(connectorType))
            {
                query += " WHERE [Type]=@Type";
                parameters.Add("Type", connectorType);
            }

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSConnectorDefinition connector = new MSConnectorDefinition
                    {
                        ConnectorId = instanceData["ConnectorId"] is DBNull ? Guid.Empty : (Guid)instanceData["ConnectorId"],
                        Name = instanceData["Name"] is DBNull ? string.Empty : (string)instanceData["Name"],
                        Type = instanceData["Type"] is DBNull ? string.Empty : (string)instanceData["Type"],
                        Icon = instanceData["Icon"] is DBNull ? string.Empty : (string)instanceData["Icon"],
                        Unique = instanceData["Unique"] is DBNull ? false : (bool)instanceData["Unique"],
                        System = instanceData["System"] is DBNull ? false : (bool)instanceData["System"],
                        Hidden = instanceData["Hidden"] is DBNull ? false : (bool)instanceData["Hidden"],
                        Parameters = instanceData["Parameters"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Parameters"]),
                    };
                    connectorDefCollection.Add(connector);
                }
            }

            return connectorDefCollection;
        }

        public List<MSConnectorDefinitionTemplate> GetConnectorDefinitionTemplates(Guid connectorDefinition, string templateType = "")
        {
            List<MSConnectorDefinitionTemplate> connectorDefTemplateCollection = new List<MSConnectorDefinitionTemplate>();
            MSPropertyBag parameters = new MSPropertyBag();

            parameters.Add("ConnectorDefinition", connectorDefinition);

            string query = sql_GetConnectorDefinitionTemplates;
            if (!string.IsNullOrWhiteSpace(templateType))
            {
                query += " AND [Type]=@Type";
                parameters.Add("Type", templateType);
            }

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSConnectorDefinitionTemplate connectorTemplate = new MSConnectorDefinitionTemplate
                    {
                        ConnectorDefinition = instanceData["ConnectorDefinition"] is DBNull ? Guid.Empty : (Guid)instanceData["ConnectorDefinition"],
                        Entity = instanceData["Entity"] is DBNull ? string.Empty : (string)instanceData["Entity"],
                        Type = instanceData["Type"] is DBNull ? string.Empty : (string)instanceData["Type"],
                        Template = instanceData["Template"] is DBNull ? string.Empty : (string)instanceData["Template"],
                        Created = instanceData["Created"] is DBNull ? DateTime.MinValue : (DateTime)instanceData["Created"],
                        CreatedBy = instanceData["CreatedBy"] is DBNull ? string.Empty : (string)instanceData["CreatedBy"],
                        Modified = instanceData["Modified"] is DBNull ? DateTime.MinValue : (DateTime)instanceData["Modified"],
                        ModifiedBy = instanceData["ModifiedBy"] is DBNull ? string.Empty : (string)instanceData["CreatedBy"]
                    };
                    connectorDefTemplateCollection.Add(connectorTemplate);
                }
            }

            return connectorDefTemplateCollection;
        }

        public List<MSRoute> GetRoute(Guid? component, string connectorType = "")
        {
            List<MSRoute> routeCollection = new List<MSRoute>();
            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetRoutes;
            string filter = string.Empty;
            if (!string.IsNullOrWhiteSpace(connectorType))
            {
                filter = " WHERE [ConnectorType]=@ConnectorType";
                parameters.Add("ConnectorType", connectorType);
            }

            if (component != null)
            {
                if (!string.IsNullOrWhiteSpace(filter))
                    filter += " AND ";
                else
                    filter += " WHERE ";

                filter += "[ComponentId]=@ComponentId";
                parameters.Add("ComponentId", component);
            }

            query += filter + " ORDER BY [Order]";

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    Guid connectorId = instanceData["Connector"] is DBNull ? Guid.Empty : (Guid)instanceData["Connector"];
                    MSConnector connector = GetConnector(connectorId);
                    MSRoute route = new MSRoute
                    {
                        Id = instanceData["RouteId"] is DBNull ? Guid.Empty : (Guid)instanceData["RouteId"],
                        Order = (int)instanceData["Order"],
                        Name = instanceData["Name"] is DBNull ? string.Empty : (string)instanceData["Name"],
                        Icon = instanceData["Icon"] is DBNull ? string.Empty : (string)instanceData["Icon"],
                        Language = instanceData["Language"] is DBNull ? null : (string)instanceData["Language"],
                        StopProcessingOnMatch = instanceData["StopProcessingOnMatch"] is DBNull ? false : (bool)instanceData["StopProcessingOnMatch"],
                        HideWorkItemLink = instanceData["HideWorkItemLink"] is DBNull ? false : (bool)instanceData["HideWorkItemLink"],
                        Conditions = instanceData["Conditions"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["Conditions"]),
                        Connector = connector,
                        ConnectorConfiguration = instanceData["ConnectorConfiguration"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["ConnectorConfiguration"]),
                        Component = instanceData["ComponentId"] is DBNull ? Guid.Empty : (Guid)instanceData["ComponentId"]
                    };
                    routeCollection.Add(route);
                }
            }

            return routeCollection;
        }

        public void UpdateRoute(MSRouteInt route)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "RouteId", route.Id },
                { "Name", route.Name },
                { "ComponentId", route.Component },
                { "Order", route.Order },
                { "Icon", route.Icon },
                { "Language", route.Language },
                { "StopProcessingOnMatch", route.StopProcessingOnMatch },
                { "HideWorkItemLink", route.HideWorkItemLink },
                { "Conditions", route.Conditions.ToString() },
                { "Connector", route.Connector },
                { "ConnectorConfiguration", route.ConnectorConfiguration.ToString() }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_UpdateRoute, parameters);
        }

        public void DeleteRoute(MSRouteInt route)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "RouteId", route.Id }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_DeleteRoute, parameters);
        }

        internal List<MSComponent>GetComponent(string query, MSPropertyBag parameters)
        {
            List<MSComponent> componentCollection = new List<MSComponent>();
            query += " ORDER BY [Name]";

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSComponent route = new MSComponent
                    {
                        Id = instanceData["ComponentId"] is DBNull ? Guid.Empty : (Guid)instanceData["ComponentId"],
                        Name = instanceData["Name"] is DBNull ? string.Empty : (string)instanceData["Name"],
                        InternalName = instanceData["InternalName"] is DBNull ? string.Empty : (string)instanceData["InternalName"],
                        Icon = instanceData["Icon"] is DBNull ? string.Empty : (string)instanceData["Icon"],
                        Capabilities = instanceData["Capabilities"] is DBNull ? null : JsonConvert.DeserializeObject<List<string>>((string)instanceData["Capabilities"]),
                        EntityProperties = instanceData["EntityProperties"] is DBNull ? null : JsonConvert.DeserializeObject((string)instanceData["EntityProperties"])
                    };
                    componentCollection.Add(route);
                }
            }

            return componentCollection;
        }

        public List<MSComponent>GetComponent(string internalName)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            string query = sql_GetComponents;

            if (!string.IsNullOrWhiteSpace(internalName))
            {
                query += " WHERE [InternalName]=@InternalName";
                parameters.Add("InternalName", internalName);
            }

            return GetComponent(query, parameters);
        }

        public List<MSComponent> GetComponent(Guid? component)
        {

            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetComponents;
            
            if (component != null)
            {
                query += " WHERE [ComponentId]=@ComponentId";
                parameters.Add("ComponentId", component);
            }

            return GetComponent(query, parameters);
        }

        public List<MSMonthlyActiveUsers> GetMAU()
        {
            List<MSMonthlyActiveUsers> mauCollection = new List<MSMonthlyActiveUsers>();
            MSPropertyBag parameters = new MSPropertyBag();

            string query = sql_GetMAU;

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
            {
                foreach (MSPropertyBag instanceData in result)
                {
                    MSMonthlyActiveUsers mau = new MSMonthlyActiveUsers
                    {
                        Service = instanceData["Service"] is DBNull ? string.Empty : (string)instanceData["Service"],
                        Active = instanceData["Active"] is DBNull ? -1 : (int)instanceData["Active"],
                        Inactive = instanceData["Inactive"] is DBNull ? -1 : (int)instanceData["Inactive"],
                        ReportDate = instanceData["ReportDate"] is DBNull ? SqlDateTime.MinValue.Value : (DateTime)instanceData["ReportDate"]
                    };
                    mauCollection.Add(mau);
                }
            }

            return mauCollection;
        }

        public MSPropertyBag GetCustomAction(Guid actionId)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            parameters.Add("ActionId", actionId);

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetCustomAction, parameters);
            if (result != null && result.Count > 0)
                return result[0];
            else
                return null;
        }

        public MSPropertyBag GetCustomAction(string internalComponentName)
        {
            List<MSComponent> components = GetComponent(internalComponentName);

            Guid componentId = Guid.Empty;

            if (components != null && components.Count > 0)
                componentId = components[0].Id;

            return componentId != Guid.Empty ? GetCustomAction(componentId) : null;
        }

        public List<MSPropertyBag> GetCustomActionsForComponent(string internalComponentName)
        {
            List<MSComponent> components = GetComponent(internalComponentName);

            if (components != null && components.Count > 0)
            {
                MSPropertyBag parameters = new MSPropertyBag();
                parameters.Add("ComponentId", components[0].Id);
                return m_sqlConn.GetDataSet(sql_GetCustomActionsForComponent, parameters);
            }
            else
                return new List<MSPropertyBag>();
        }

        public List<MSPropertyBag> GetCustomActionsForComponent(Guid component)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            parameters.Add("ComponentId", component);
            return m_sqlConn.GetDataSet(sql_GetCustomActionsForComponent, parameters);
        }

        public void UpdateCustomAction(CustomBaseActionInt action)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ActionId", action.id },              
                { "ComponentId", action.component },
                { "Name", action.name },
                { "Icon", action.icon },
                { "Type", action.type },
                { "Configuration", action.configuration.ToString() }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_AddCustomAction, parameters);
        }

        public void DeleteCustomAction(CustomBaseActionInt action)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ActionId", action.id }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_DeleteCustomAction, parameters);
        }

        public MSPropertyBag GetCustomActionType(Guid typeId)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            parameters.Add("ActionTypeId", typeId);

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetCustomActionType, parameters);
            if (result != null && result.Count > 0)
                return result[0];
            else
                return null;
        }

        public List<MSPropertyBag> GetCustomActionTypes()
        {
            MSPropertyBag parameters = new MSPropertyBag();

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetCustomActionTypes, parameters);
            if (result != null && result.Count > 0)
                return result;
            else
                return null;
        }

        public List<MSPropertyBag> GetImages(string type)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            string query = sql_GetImages;

            if (!string.IsNullOrWhiteSpace(type))
            {
                query += " WHERE [Type]=@Type";
                parameters.Add("Type", type);
            }

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
                return result;
            else
                return new List<MSPropertyBag>();
        }

        public List<MSPropertyBag> GetImage(string name, string type)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            string query = sql_GetImages + " WHERE [Name]=@Name AND [Type]=@Type"; ;
            parameters.Add("Type", type);
            parameters.Add("Name", name);

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
                return result;
            else
                return new List<MSPropertyBag>();
        }

        public List<MSPropertyBag> GetAuditLog(string item)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            string query = sql_GetAuditLog;

            query += " WHERE [Item]=@Item ORDER BY [Timestamp] DESC";
            parameters.Add("Item", item);

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
                return result;
            else
                return new List<MSPropertyBag>();
        }

        public List<MSPropertyBag> GetAuditLog(Guid correlationId)
        {
            MSPropertyBag parameters = new MSPropertyBag();
            string query = sql_GetAuditLog;

            query += " WHERE [CorrelationId]=@CorrelationId ORDER BY [Timestamp] DESC";
            parameters.Add("CorrelationId", @correlationId);

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(query, parameters);
            if (result != null && result.Count > 0)
                return result;
            else
                return new List<MSPropertyBag>();
        }

        public void AddAuditLogRecord(
            Guid organizationId,
            Guid correlationId,
            string ipAddress,
            string user,
            string activity,
            string item,
            string itemType,
            Guid versionId,
            string itemUniqueId,
            string eventSource,
            Guid componentId,
            dynamic extendedProperties,
            dynamic modifiedProperties)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "OrganizationId", organizationId },
                { "Activity", activity },
                { "Item", item },
                { "ItemType", itemType },
                { "ItemUniqueId", itemUniqueId },
                { "EventSource", eventSource }
            };

            if (correlationId != Guid.Empty)
                parameters.Add("CorrelationId", correlationId);

            if(!string.IsNullOrWhiteSpace(ipAddress))
                parameters.Add("IPAddress", ipAddress);

            if (!string.IsNullOrWhiteSpace(user))
                parameters.Add("User", user);

            if (versionId != Guid.Empty)
                parameters.Add("VersionId", versionId);

            if (componentId != Guid.Empty)
                parameters.Add("ComponentId", componentId);

            if (extendedProperties != null)
                parameters.Add("ExtendedProperties", JsonConvert.SerializeObject(extendedProperties));

            if (modifiedProperties != null)
                parameters.Add("ModifiedProperties", JsonConvert.SerializeObject(modifiedProperties));

            m_sqlConn.ExecuteStoredProcedure(sql_sp_AddActivityLogRecord, parameters);
        }

        public int SetItemArchiveFlag(string id, string type, bool archived, string userName)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "Id", id },
                { "Type", type },
                { "Archived", archived }
            };

            int res = m_sqlConn.ExecuteStoredProcedure(sql_sp_SetArchiveFlag, parameters);

            if (res == 1)
            {
                AddAuditLogRecord(Guid.Empty, Guid.Empty, "", userName, archived ? "Archived" : "Restored",
                            string.Format("item://{0}/{1}", type, id), "Item",
                            Guid.Empty, id, "WebApp", Guid.Empty, null, null);
            }

            Cache.Instance.ReloadPersistedContent(id, type);

            return res;
        }

        public List<MSPropertyBag> GetUserProfile(Guid objectId)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ObjectId", objectId }
            };

            List<MSPropertyBag> result = m_sqlConn.GetDataSet(sql_GetUserProfile, parameters);

            return result;
        }

        public int SetUserProfileProperties(Guid objectId, string propertiesJson)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "ObjectId", objectId },
                { "PropertiesJson", propertiesJson }
            };

            int res = m_sqlConn.ExecuteStoredProcedure(sql_sp_SetUserProfileProperties, parameters);

            return res;
        }

        public List<MSPropertyBag> GetJobStatistics(Guid id, int top, bool? hideEmptyItems, DateTime? startDate, DateTime? endDate)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "JobId", id }
            };

            string query = string.Empty;

            if (top > 0)
            {
                parameters.Add("ItemCount", top);
                query = sql_GetTopJobStatistics;
                
            } else
            {
                query = sql_GetJobStatistics;
            }

            query += hideEmptyItems != null && hideEmptyItems.Value ? " AND ([ItemsCreated] > 0 OR [ItemsModified] > 0 OR [ItemsFailed] > 0 OR [TasksCreated] > 0 OR [TasksModified] > 0 OR [TasksFailed] > 0 OR [NotificationsSent] > 0 OR [NotificationsFailed] > 0)" : "";

            if (startDate != null)
            {
                query += " AND [Start]>=@StartDate";
                parameters.Add("StartDate", startDate.Value);
            }

            if (endDate != null)
            {
                query += " AND [End]<=@EndDate";
                parameters.Add("EndDate", endDate.Value);
            }

            query += " ORDER BY [js].[Start] DESC";
            result = m_sqlConn.GetDataSet(query, parameters);
            return result;
        }

        public List<MSPropertyBag> GetJobStatistics(Guid id, string communicationId)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "JobId", id },
                { "ItemId", communicationId + "%" }
            };

            result = m_sqlConn.GetDataSet(sql_GetJobStatisticsForComm, parameters);

            return result;
        }

        public bool GetJobTimeoutState(Guid correlationId, int timeoutThreshold = 10)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "CorrelationId", correlationId },
                { "TimeoutThreshold", timeoutThreshold }
            };

            result = m_sqlConn.GetDataSet(sql_GetJobTimeoutState, parameters);

            if (result.Count > 0)
            {
                bool timeoutState = result[0]["TimeoutThresholdReached"] is DBNull ? false : ((int)result[0]["TimeoutThresholdReached"])>0;
                return timeoutState;
            }
            else
                return false;
        }

        public List<MSPropertyBag> GetJobStatistics(Guid id)
        {
            return GetJobStatistics(id, 0, null, null, null);
        }

        public List<MSPropertyBag> GetMessageCenterWeeklyStatistics(string messageType)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "messageType", messageType }
            };
            
            result = m_sqlConn.GetDataSet(sql_GetWeeklyMCStatistics, parameters);
            return result;
        }

        public List<MSPropertyBag> GetMessageCenterPastWeekStatistics()
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new();

            result = m_sqlConn.GetDataSet(sql_GetMCStatisticsForPast7Days, parameters);
            return result;
        }

        public List<MSPropertyBag> GetAzureHealthAlerts(bool? active=null, Guid? userId=null)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new ();

            string query = sql_GetAzureServiceHealthAlerts;

            if (active != null)
            {
                query += " AND JSON_VALUE([Data], '$.data.alertContext.status')=@isActive";
                parameters.Add("isActive", active.Value ? "Active" : "Resolved");
            }

            result = m_sqlConn.GetDataSet(query, parameters);

            List<string> ids = new();
            MSCommunicationIdCollection idCollection = new MSCommunicationIdCollection();

            foreach (MSPropertyBag record in result) {
                if (record["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)record["ID"]))
                {
                    string id = (string)record["ID"];
                    if (!ids.Contains(id.Trim()))
                    {
                        ids.Add(id.Trim());

                        if (userId != null && userId != Guid.Empty)
                            idCollection.Add(new MSCommunicationId()
                            {
                                userId = userId.Value,
                                communicationId = id.Trim(),
                                tenantId = Guid.Empty,        // tenantId and subscriptionId shall be adjusted
                                subscriptionId = Guid.Empty   // during multitenancy implementation
                            });
                    }
                }              
            }

            if (ids.Count> 0)
            {
                List<MSSummaryCacheEntry> summaryList = GetLastNotificationSummary(ids);
                List<MSViewpoint> viewpoints = GetViewpoints(idCollection);
                List<MSCommunicationTag> tags = GetTagCollection(idCollection);

                foreach (MSPropertyBag rec in result)
                {
                    if (rec["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)rec["ID"]))
                    {
                        string id = (string)rec["ID"];
                        MSSummaryCacheEntry? summary = summaryList?.Find(s => s.id == id);
                        dynamic additionalData = new ExpandoObject();
                        additionalData.summary = summary;

                        if (rec.ContainsKey("additionalData"))
                            rec["additionalData"] = additionalData;
                        else
                            rec.Add("additionalData", additionalData);

                        dynamic serviceHealthHubViewpoint = new ExpandoObject();

                        MSViewpoint viewpoint = viewpoints.Find(v => v.communicationId == id); // add check for tenantId and subscriptionId
                        serviceHealthHubViewpoint.archived = viewpoint != null && viewpoint.archived.Value;
                        serviceHealthHubViewpoint.favorited = viewpoint != null && viewpoint.favorite.Value;
                        serviceHealthHubViewpoint.viewed = viewpoint != null && viewpoint.read.Value;

                        if (rec.ContainsKey("serviceHealthHubViewpoint"))
                            rec["serviceHealthHubViewpoint"] = serviceHealthHubViewpoint;
                        else
                            rec.Add("serviceHealthHubViewpoint", serviceHealthHubViewpoint);

                        List<MSCommunicationTag> commTags = tags.FindAll(t => t.MessageId == id && t.TagId != null);
                        if (rec.ContainsKey("organizationTags"))
                            rec["organizationTags"] = commTags;
                        else
                            rec.Add("organizationTags", commTags);
                    }
                }
            }

            return result;
        }

        public MSPropertyBag GetAzureHealthAlert(string id, Guid? userId = null)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new();

            string query = sql_GetAzureServiceHealthAlerts;
            query += " AND [ID]=@Id";
            parameters.Add("Id", id);

            result = m_sqlConn.GetDataSet(query, parameters);

            if (result == null || result.Count <= 0)
                return null;

            List<string> ids = new();
            MSPropertyBag record = result[0];

            MSCommunicationIdCollection idCollection = new();
            if (userId != null && userId != Guid.Empty)
                idCollection.Add(new MSCommunicationId()
                {
                    userId = userId.Value,
                    communicationId = id.Trim(),
                    tenantId = Guid.Empty,        // tenantId and subscriptionId shall be adjusted
                    subscriptionId = Guid.Empty   // during multitenancy implementation
                });

            if (record["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)record["ID"]))
            {
                ids.Add((string)record["ID"]);
                List<MSSummaryCacheEntry> summaryList = GetLastNotificationSummary(ids);
                List<MSViewpoint> viewpoints = GetViewpoints(idCollection);
                List<MSCommunicationTag> tags = GetTagCollection(idCollection);

                MSSummaryCacheEntry? summary = summaryList?.Find(s => s.id == (string)record["ID"]);
                dynamic additionalData = new ExpandoObject();
                additionalData.summary = summary;

                if (record.ContainsKey("additionalData"))
                    record["additionalData"] = additionalData;
                else
                    record.Add("additionalData", additionalData);

                dynamic serviceHealthHubViewpoint = new ExpandoObject();

                MSViewpoint viewpoint = viewpoints.Find(v => v.communicationId == id); // add check for tenantId and subscriptionId
                serviceHealthHubViewpoint.archived = viewpoint != null && viewpoint.archived.Value;
                serviceHealthHubViewpoint.favorited = viewpoint != null && viewpoint.favorite.Value;
                serviceHealthHubViewpoint.viewed = viewpoint != null && viewpoint.read.Value;

                if (record.ContainsKey("serviceHealthHubViewpoint"))
                    record["serviceHealthHubViewpoint"] = serviceHealthHubViewpoint;
                else
                    record.Add("serviceHealthHubViewpoint", serviceHealthHubViewpoint);

                List<MSCommunicationTag> commTags = tags.FindAll(t => t.MessageId == id && t.TagId != null);
                if (record.ContainsKey("organizationTags"))
                    record["organizationTags"] = commTags;
                else
                    record.Add("organizationTags", commTags);
            }

            return record;
        }

        public List<MSPropertyBag> GetAzureUpdates(bool? active = null, Guid? userId = null)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new();

            string query = sql_GetAzureUpdates;

            if (active != null)
            {
                if (active == true)
                    query = sql_GetAzureUpdatesActive;
                else
                {
                    query += " AND JSON_VALUE([Data], '$.releaseStatus')" + (active.Value ? " <> " : " = ") + "'Launched'";
                }
            }

            result = m_sqlConn.GetDataSet(query, parameters);

            List<string> ids = new();
            MSCommunicationIdCollection idCollection = new MSCommunicationIdCollection();

            foreach (MSPropertyBag record in result)
            {
                if (record["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)record["ID"]))
                {
                    string id = (string)record["ID"];
                    if (!ids.Contains(id.Trim()))
                    {
                        ids.Add(id.Trim());

                        if (userId != null && userId != Guid.Empty)
                            idCollection.Add(new MSCommunicationId()
                            {
                                userId = userId.Value,
                                communicationId = id.Trim(),
                                tenantId = Guid.Empty,        // tenantId and subscriptionId shall be adjusted
                                subscriptionId = Guid.Empty   // during multitenancy implementation
                            });
                    }
                }
            }

            if (ids.Count > 0)
            {
                List<MSSummaryCacheEntry> summaryList = GetLastNotificationSummary(ids);
                List<MSViewpoint> viewpoints = GetViewpoints(idCollection);
                List<MSCommunicationTag> tags = GetTagCollection(idCollection);

                foreach (MSPropertyBag rec in result)
                {
                    if (rec["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)rec["ID"]))
                    {
                        string id = (string)rec["ID"];
                        MSSummaryCacheEntry? summary = summaryList?.Find(s => s.id == id);
                        dynamic additionalData = new ExpandoObject();
                        additionalData.summary = summary;

                        if (rec.ContainsKey("additionalData"))
                            rec["additionalData"] = additionalData;
                        else
                            rec.Add("additionalData", additionalData);

                        dynamic serviceHealthHubViewpoint = new ExpandoObject();

                        MSViewpoint viewpoint = viewpoints.Find(v => v.communicationId == id); // add check for tenantId and subscriptionId
                        serviceHealthHubViewpoint.archived = viewpoint != null && viewpoint.archived.Value;
                        serviceHealthHubViewpoint.favorited = viewpoint != null && viewpoint.favorite.Value;
                        serviceHealthHubViewpoint.viewed = viewpoint != null && viewpoint.read.Value;

                        if (rec.ContainsKey("serviceHealthHubViewpoint"))
                            rec["serviceHealthHubViewpoint"] = serviceHealthHubViewpoint;
                        else
                            rec.Add("serviceHealthHubViewpoint", serviceHealthHubViewpoint);

                        List<MSCommunicationTag> commTags = tags.FindAll(t => t.MessageId == id && t.TagId != null);
                        if (rec.ContainsKey("organizationTags"))
                            rec["organizationTags"] = commTags;
                        else
                            rec.Add("organizationTags", commTags);
                    }
                }
            }

            return result;
        }

        public MSPropertyBag GetAzureUpdate(string id, Guid? userId = null)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new();

            string query = sql_GetAzureUpdates;
            query += " AND [ID]=@Id";
            parameters.Add("Id", id);

            result = m_sqlConn.GetDataSet(query, parameters);

            if (result == null || result.Count <= 0)
                return null;

            List<string> ids = new();
            MSPropertyBag record = result[0];

            MSCommunicationIdCollection idCollection = new();
            if (userId != null && userId != Guid.Empty)
                idCollection.Add(new MSCommunicationId()
                {
                    userId = userId.Value,
                    communicationId = id.Trim(),
                    tenantId = Guid.Empty,        // tenantId and subscriptionId shall be adjusted
                    subscriptionId = Guid.Empty   // during multitenancy implementation
                });

            if (record["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)record["ID"]))
            {
                ids.Add((string)record["ID"]);
                List<MSSummaryCacheEntry> summaryList = GetLastNotificationSummary(ids);
                List<MSViewpoint> viewpoints = GetViewpoints(idCollection);
                List<MSCommunicationTag> tags = GetTagCollection(idCollection);

                MSSummaryCacheEntry? summary = summaryList?.Find(s => s.id == (string)record["ID"]);
                dynamic additionalData = new ExpandoObject();
                additionalData.summary = summary;

                if (record.ContainsKey("additionalData"))
                    record["additionalData"] = additionalData;
                else
                    record.Add("additionalData", additionalData);

                dynamic serviceHealthHubViewpoint = new ExpandoObject();

                MSViewpoint viewpoint = viewpoints.Find(v => v.communicationId == id); // add check for tenantId and subscriptionId
                serviceHealthHubViewpoint.archived = viewpoint != null && viewpoint.archived.Value;
                serviceHealthHubViewpoint.favorited = viewpoint != null && viewpoint.favorite.Value;
                serviceHealthHubViewpoint.viewed = viewpoint != null && viewpoint.read.Value;

                if (record.ContainsKey("serviceHealthHubViewpoint"))
                    record["serviceHealthHubViewpoint"] = serviceHealthHubViewpoint;
                else
                    record.Add("serviceHealthHubViewpoint", serviceHealthHubViewpoint);

                List<MSCommunicationTag> commTags = tags.FindAll(t => t.MessageId == id && t.TagId != null);
                if (record.ContainsKey("organizationTags"))
                    record["organizationTags"] = commTags;
                else
                    record.Add("organizationTags", commTags);
            }

            return record;
        }

        public List<MSPropertyBag> GetD365PowerPlatformReleases(Guid? userId = null)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new();

            string query = sql_GetD365PowerPlatformReleases;
            query += " WHERE [GADate] > DATEADD(month, -6, GETUTCDATE()) OR [GADate] IS NULL";

            result = m_sqlConn.GetDataSet(query, parameters);

            List<string> ids = new();
            MSCommunicationIdCollection idCollection = new MSCommunicationIdCollection();

            foreach (MSPropertyBag record in result)
            {
                if (record["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)record["ID"]))
                {
                    string id = (string)record["ID"];
                    if (!ids.Contains(id.Trim()))
                    {
                        ids.Add(id.Trim());

                        if (userId != null && userId != Guid.Empty)
                            idCollection.Add(new MSCommunicationId()
                            {
                                userId = userId.Value,
                                communicationId = id.Trim(),
                                tenantId = Guid.Empty,        // tenantId and subscriptionId shall be adjusted
                                subscriptionId = Guid.Empty   // during multitenancy implementation
                            });
                    }
                }
            }

            if (ids.Count > 0)
            {
                List<MSViewpoint> viewpoints = GetViewpoints(idCollection);
                List<MSCommunicationTag> tags = GetTagCollection(idCollection);

                foreach (MSPropertyBag rec in result)
                {
                    if (rec["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)rec["ID"]))
                    {
                        string id = (string)rec["ID"];

                        dynamic serviceHealthHubViewpoint = new ExpandoObject();

                        MSViewpoint viewpoint = viewpoints.Find(v => v.communicationId == id); // add check for tenantId and subscriptionId
                        serviceHealthHubViewpoint.archived = viewpoint != null && viewpoint.archived.Value;
                        serviceHealthHubViewpoint.favorited = viewpoint != null && viewpoint.favorite.Value;
                        serviceHealthHubViewpoint.viewed = viewpoint != null && viewpoint.read.Value;

                        if (rec.ContainsKey("serviceHealthHubViewpoint"))
                            rec["serviceHealthHubViewpoint"] = serviceHealthHubViewpoint;
                        else
                            rec.Add("serviceHealthHubViewpoint", serviceHealthHubViewpoint);

                        List<MSCommunicationTag> commTags = tags.FindAll(t => t.MessageId == id && t.TagId != null);
                        if (rec.ContainsKey("organizationTags"))
                            rec["organizationTags"] = commTags;
                        else
                            rec.Add("organizationTags", commTags);
                    }
                }
            }

            return result;
        }

        public MSPropertyBag? GetD365PowerPlatformRelease(string id, Guid? userId = null)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new();

            string query = sql_GetD365PowerPlatformReleases;
            query += " WHERE [ID]=@Id";
            parameters.Add("Id", id);

            result = m_sqlConn.GetDataSet(query, parameters);

            if (result == null || result.Count <= 0)
                return null;

            List<string> ids = new();
            MSPropertyBag record = result[0];

            MSCommunicationIdCollection idCollection = new();
            if (userId != null && userId != Guid.Empty)
                idCollection.Add(new MSCommunicationId()
                {
                    userId = userId.Value,
                    communicationId = id.Trim(),
                    tenantId = Guid.Empty,        // tenantId and subscriptionId shall be adjusted
                    subscriptionId = Guid.Empty   // during multitenancy implementation
                });

            if (record["ID"] is not DBNull && !string.IsNullOrWhiteSpace((string)record["ID"]))
            {
                ids.Add((string)record["ID"]);
                List<MSViewpoint> viewpoints = GetViewpoints(idCollection);
                List<MSCommunicationTag> tags = GetTagCollection(idCollection);

                dynamic serviceHealthHubViewpoint = new ExpandoObject();

                MSViewpoint viewpoint = viewpoints.Find(v => v.communicationId == id); // add check for tenantId and subscriptionId
                serviceHealthHubViewpoint.archived = viewpoint != null && viewpoint.archived.Value;
                serviceHealthHubViewpoint.favorited = viewpoint != null && viewpoint.favorite.Value;
                serviceHealthHubViewpoint.viewed = viewpoint != null && viewpoint.read.Value;

                if (record.ContainsKey("serviceHealthHubViewpoint"))
                    record["serviceHealthHubViewpoint"] = serviceHealthHubViewpoint;
                else
                    record.Add("serviceHealthHubViewpoint", serviceHealthHubViewpoint);

                List<MSCommunicationTag> commTags = tags.FindAll(t => t.MessageId == id && t.TagId != null);
                if (record.ContainsKey("organizationTags"))
                    record["organizationTags"] = commTags;
                else
                    record.Add("organizationTags", commTags);
            }

            return record;
        }

        public List<MSPropertyBag> GetMonthlyActiveUsersReport(string[]? services)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new();

            string query = sql_GetMonthlyActiveUsersReport;

            if (services != null)
            {
                List<string> svcList = new();
                foreach (string service in services)
                    svcList.Add(service.Trim());

                string svcListJoined = "'" + string.Join("','", svcList) + "'";
                query += " WHERE [Service] IN (" + svcListJoined + ")";
            }

            result = m_sqlConn.GetDataSet(query, parameters);
            return result;
        }

        public MSSummaryCacheEntry GetLastNotificationSummary(string id)
        {
            List<MSPropertyBag> result;
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "MessageId", id }
            };

            result = m_sqlConn.GetDataSet(sql_GetLastNotificationSummary, parameters);

            if (result == null || result.Count <= 0)
                return null;
            else
            {
                MSSummaryCacheEntry returnObject = new MSSummaryCacheEntry();
                MSPropertyBag summaryRecord = result[0];

                if (summaryRecord["Timestamp"] is DBNull)
                    returnObject.timestamp = null;
                else
                    returnObject.timestamp = (DateTime)summaryRecord["Timestamp"];

                if (summaryRecord["Summary"] is DBNull)
                    returnObject.contents = null;
                else
                    returnObject.contents = JsonConvert.DeserializeObject<object>((string)summaryRecord["Summary"]);

                return returnObject;
            }
        }

        public List<MSSummaryCacheEntry> GetLastNotificationSummary(List<string> id)
        {
            if (id == null || id.Count <= 0) 
                return null;
            
            List<MSPropertyBag> result;

            string idList = "'" + string.Join("','", id.ToArray()) + "'";
            string query = sql_GetLastNotificationSummaryList + string.Format(" WHERE [MessageId] IN ({0})", idList);
            MSPropertyBag parameters = new();

            result = m_sqlConn.GetDataSet(query, parameters);

            if (result == null || result.Count <= 0)
                return null;
            else
            {
                List<MSSummaryCacheEntry> summaryList = new();
                foreach (MSPropertyBag summaryRecord in result)
                {
                    if (summaryRecord["MessageId"] is not DBNull)
                    {
                        MSSummaryCacheEntry returnObject = new();

                        returnObject.id = (string)summaryRecord["MessageId"];

                        if (summaryRecord["Timestamp"] is DBNull)
                            returnObject.timestamp = null;
                        else
                            returnObject.timestamp = (DateTime)summaryRecord["Timestamp"];

                        if (summaryRecord["Summary"] is DBNull)
                            returnObject.contents = null;
                        else
                            returnObject.contents = JsonConvert.DeserializeObject<object>((string)summaryRecord["Summary"]);

                        summaryList.Add(returnObject);
                    }
                }

                return summaryList;
            }
        }

        public List<MSViewpoint> GetViewpoints(MSCommunicationIdCollection communicationIds)
        {
            List<MSViewpoint> viewpointList = new List<MSViewpoint>();

            if (communicationIds == null || communicationIds.Count <= 0)
                return viewpointList;

            List<MSPropertyBag> result;
            MSPropertyBag parameters = new()
            {
                { 
                    "CommunicationIds", new MSSqlStructuredParameter()
                    {
                        TypeName = "dbo.tvp_CommunicationIdCollection",
                        Value = communicationIds
                    }
                }
            };

            result = m_sqlConn.GetDataSet(sql_GetViewpoint, parameters);

            if (result == null || result.Count <= 0)
                return viewpointList;
            else
            {
                foreach (MSPropertyBag record in result)
                {
                    MSViewpoint viewpoint = new MSViewpoint()
                    {
                        userId = !record.ContainsKey("UserId") || record["UserId"] is DBNull ? Guid.Empty : (Guid)record["UserId"],
                        communicationId = !record.ContainsKey("CommunicationId") || record["CommunicationId"] is DBNull ? string.Empty : (string)record["CommunicationId"],
                        tenantId = !record.ContainsKey("TenantId") || record["TenantId"] is DBNull ? Guid.Empty : (Guid)record["TenantId"],
                        subscriptionId = !record.ContainsKey("SubscriptionId") || record["SubscriptionId"] is DBNull ? Guid.Empty : (Guid)record["SubscriptionId"],
                        read = !record.ContainsKey("Read") || record["Read"] is DBNull ? false : (bool)record["Read"],
                        favorite = !record.ContainsKey("Favorite") || record["Favorite"] is DBNull ? false : (bool)record["Favorite"],
                        archived = !record.ContainsKey("Archived") || record["Archived"] is DBNull ? false : (bool)record["Archived"]
                    };
                    
                    viewpointList.Add(viewpoint);
                }

                return viewpointList;
            }
        }

        public async void SetViewpointReadFlag(Guid userId, string communicationId, Guid tenantId, Guid subscriptionId, bool viewed)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "UserId", userId },
                { "CommunicationId", communicationId },
                { "TenantId", tenantId },
                { "SubscriptionId", subscriptionId },
                { "Read", viewed }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_SetViewpointReadFlag, parameters);

            // if Graph Connector is active and enabled, register interaction event.
            try
            {
                MSSearchSettings search = new MSSearchSettings();
                if (search.ConnectorSettings.Configured && search.GraphAPIConnector != null)
                {
                    AddActivitiesPostRequestBody body = new AddActivitiesPostRequestBody();
                    ExternalActivity activity = new ExternalActivity();
                    var identity = new Microsoft.Graph.Models.ExternalConnectors.Identity();
                    identity.Type = IdentityType.User;
                    identity.Id = userId.ToString();
                    activity.Type = ExternalActivityType.Viewed;
                    activity.PerformedBy = identity;
                    activity.StartDateTime = DateTime.UtcNow;
                    body.Activities = new List<ExternalActivity> { activity };

                    var result = await GraphApiClientHelper.Client.External.Connections[search.GraphAPIConnector.Id].Items[communicationId].MicrosoftGraphExternalConnectorsAddActivities.PostAsAddActivitiesPostResponseAsync(body)
                        .ContinueWith(t =>
                        {
                            return t.Result;
                        });
                }
            } catch
            {
                
            }
        }

        public void SetViewpointArchiveFlag(Guid userId, string communicationId, Guid tenantId, Guid subscriptionId, bool archived)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "UserId", userId },
                { "CommunicationId", communicationId },
                { "TenantId", tenantId },
                { "SubscriptionId", subscriptionId },
                { "Archived", archived }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_SetViewpointArchiveFlag, parameters);
        }

        public void SetViewpointFavoriteFlag(Guid userId, string communicationId, Guid tenantId, Guid subscriptionId, bool favorite)
        {
            MSPropertyBag parameters = new MSPropertyBag
            {
                { "UserId", userId },
                { "CommunicationId", communicationId },
                { "TenantId", tenantId },
                { "SubscriptionId", subscriptionId },
                { "Favorite", favorite}
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_SetViewpointFavoriteFlag, parameters);
        }

        #region User-defined tags
        public List<MSPropertyBag> GetTagDefinitions()
        {
            return m_sqlConn.GetDataSet(sql_GetTagDefinitions);
        }

        public List<MSPropertyBag> GetTagDefinition(Guid tagId)
        {
            MSPropertyBag parameters = new()
            {
                { "TagId", tagId }
            };

            return m_sqlConn.GetDataSet(sql_GetTagDefinition, parameters);
        }

        public List<MSPropertyBag> GetTags(string messageId, string type)
        {
            MSPropertyBag parameters = new()
            {
                { "MessageId", messageId },
                { "Type", type }
            };

            return m_sqlConn.GetDataSet(sql_GetTags, parameters);
        }

        public void MoveTagDefinition(Guid tagId, string? targetGroup)
        {
            MSPropertyBag parameters = new()
            {
                { "TagId", tagId },
                { "Type", targetGroup }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_MoveTagDefinition, parameters);
        }

        public void AddTagDefinition(Guid tagId, string name, string? targetGroup)
        {
            MSPropertyBag parameters = new()
            {
                { "TagId", tagId },
                { "Name", name },
                { "Type", targetGroup }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_AddTagDefinition, parameters);
        }

        public void RemoveTagDefinition(Guid tagId)
        {
            MSPropertyBag parameters = new()
            {
                { "TagId", tagId }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_RemoveTagDefinition, parameters);
        }

        public List<MSPropertyBag> GetTag(string messageId, string type, Guid tagId)
        {
            MSPropertyBag parameters = new()
            {
                { "MessageId", messageId },
                { "Type", type },
                { "TagId", tagId }
            };

            return m_sqlConn.GetDataSet(sql_GetTags + " AND [TagId] = @TagId", parameters);
        }

        public void AddTag(string messageId, string type, Guid tagId)
        {
            MSPropertyBag parameters = new()
            {      
                { "MessageId", messageId },
                { "Type", type },
                { "TagId", tagId }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_AddTag, parameters);
        }

        public void RemoveTag(string messageId, string type, Guid tagId)
        {
            MSPropertyBag parameters = new()
            {
                { "MessageId", messageId },
                { "Type", type },
                { "TagId", tagId }
            };

            m_sqlConn.ExecuteStoredProcedure(sql_sp_RemoveTag, parameters);
        }

        public List<MSCommunicationTag> GetTagCollection(MSCommunicationIdCollection communicationIds)
        {
            List<MSCommunicationTag> tagList = new List<MSCommunicationTag>();

            if (communicationIds == null || communicationIds.Count <= 0)
                return tagList;

            List<MSPropertyBag> result;
            MSPropertyBag parameters = new()
            {
                {
                    "CommunicationIds", new MSSqlStructuredParameter()
                    {
                        TypeName = "dbo.tvp_CommunicationIdCollection",
                        Value = communicationIds
                    }
                }
            };

            result = m_sqlConn.GetDataSet(sql_GetTagCollection, parameters);

            if (result == null || result.Count <= 0)
                return tagList;
            else
            {
                foreach (MSPropertyBag record in result)
                {
                    MSCommunicationTag tag = MSCommunicationTag.CreateInstance(record);

                    tagList.Add(tag);
                }

                return tagList;
            }
        }
        #endregion
    }
}
