using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Primitives;
using Microsoft.Graph.Models;
using Microsoft.ServiceHealthHub.Graph;
using System.Text.RegularExpressions;

namespace Microsoft.ServiceHealthHub.Core
{
    public sealed class Cache
    {
        const string c_serviceHealthCacheName = "ServiceHealthCollection-ce2cda9c-a8ec-4ac7-a2c4-af5aba524e6e";
        const string c_messageCenterCacheName = "MessageCenterCollection-a8ce2447-3a11-4479-bc04-91e80d2a3fe6";
        const string c_serviceHealthOverviewCacheName = "ServiceHealthOverviewCollection-e8b5d354-b8bf-4a44-ba19-32c3a51a6e84";

        IMemoryCache m_cache = new MemoryCache(new MemoryCacheOptions());
        private static readonly Cache m_Instance = new Cache();

        public static Cache Instance
        {
            get
            {
                return m_Instance;
            }
        }

        // Explicit static constructor to tell C# compiler
        // not to mark type as beforefieldinit
        static Cache()
        {
        }

        private Cache()
        {
            
        }

        public bool Exists<T>(string key) where T : class
        {
            T value;
            return m_cache.TryGetValue(key, out value);
        }

        public void ReloadPersistedContent(string id, string type)
        {
            MSSHNotificationDatabase db = new();
            MSServiceNotification record = new();
            db.GetNotification(id, record);
            MSSummaryCacheEntry summary = db.GetLastNotificationSummary(id);

            if (record != null && !string.IsNullOrWhiteSpace(record.Id))
            {
                switch (type.ToLower())
                {
                    case "servicehealthissue":
                        if (Exists<List<M365ServiceIssue>>(c_serviceHealthCacheName))
                        {
                            List<M365ServiceIssue> issues = GetServiceHealthIssueCollection();
                            M365ServiceIssue cacheEntry = issues?.Find(i => i.Id == record.Id);
                            if (cacheEntry != null)
                                RehydratePersistedContent(cacheEntry, record, summary);
                        }
                        break;
                    case "serviceupdatemessage":
                        if (Exists<List<M365MessageCenterItem>>(c_messageCenterCacheName))
                        {
                            List<M365MessageCenterItem> messages = GetMessageCenterCollection();
                            M365MessageCenterItem cacheEntry = messages?.Find(m => m.Id == record.Id);
                            if (cacheEntry != null)
                                RehydratePersistedMessageCenterContent(cacheEntry, record, summary);
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        private void RehydratePersistedContent(M365ServiceIssue issue, MSServiceNotification persistedContent, MSSummaryCacheEntry summary)
        {
            try
            {
                issue.Public = persistedContent.Public;
                issue.PublishingComments = persistedContent.Comments;
                issue.ExtendedProperties = persistedContent.ExtendedProperties;
                issue.Summary = summary;
                issue.ServiceHealthHubState = persistedContent.State;

                if (!string.IsNullOrWhiteSpace(persistedContent.WorkItemID))
                {
                    issue.Task.TaskId = persistedContent.WorkItemID;
                    issue.Task.TaskUrl = persistedContent.WorkItemURL;
                }
            }
            catch
            {

            }
        }

        private List<M365ServiceIssue> GetServiceHealthIssuesFromAPI()
        {
            List<ServiceHealthIssue> allIssues = new List<ServiceHealthIssue>();
            List<M365ServiceIssue> result = new List<M365ServiceIssue>();

            try
            {

                var apiCall = GraphApiClientHelper.Client.Admin.ServiceAnnouncement.Issues;

                Task<ServiceHealthIssueCollectionResponse?> t = Task.Run(async () =>
                {
                    ServiceHealthIssueCollectionResponse? res = await apiCall.GetAsync();
                    return res;
                });

                t.Wait();

                ServiceHealthIssueCollectionResponse? issues = t.Result;
                
                allIssues.AddRange(issues?.Value);

                while (issues?.OdataNextLink != null)
                {
                    t = Task.Run(async () =>
                    {
                        ServiceHealthIssueCollectionResponse? res = await apiCall.WithUrl(issues?.OdataNextLink).GetAsync();
                        return res;
                    });
                    t.Wait();
                    issues = t.Result;

                    allIssues.AddRange(issues?.Value);
                }

                List<string> idList = new List<string>();
                foreach (ServiceHealthIssue i in allIssues)
                {
                    if (i.Id != null)
                        idList.Add(i.Id);
                }

                MSSHNotificationDatabase db = new();
                List<MSServiceNotification> dbNotifications = new List<MSServiceNotification>();

                foreach (var apiIssue in allIssues)
                {
                    M365ServiceIssue issue = M365ServiceIssue.From(apiIssue);

                    if (issue.AdditionalData == null)
                        issue.AdditionalData = new Dictionary<string, object>();

                    if (null != issue.Status)
                    {
                        Dictionary<string, string> statusDisplayNames = new()
                    {
                        { "serviceoperational", "Service operational" },
                        { "investigating", "Investigating" },
                        { "restoringservice", "Restoring service" },
                        { "verifyingservice", "Verifying service" },
                        { "servicerestored", "Service restored" },
                        { "postincidentreviewpublished", "Post-incident review published" },
                        { "servicedegradation", "Service degradation" },
                        { "serviceinterruption", "Service interruption" },
                        { "extendedrecovery", "Extended recovery" },
                        { "falsepositive", "False positive" },
                        { "investigationsuspended", "Investigation suspended" },
                        { "resolved", "Resolved" },
                        { "mitigatedexternal", "Mitigated (external)" },
                        { "mitigated", "Mitigated" },
                        { "resolvedexternal", "Resolved (external)" },
                        { "confirmed", "Confirmed" },
                        { "reported", "Reported" }
                    };

                        if (statusDisplayNames.ContainsKey(issue.Status.ToString().ToLower()))
                            issue.StatusDisplayName = statusDisplayNames[issue.Status.ToString().ToLower()];
                        else
                            issue.StatusDisplayName = Regex.Replace(issue.Status.ToString(), "([a-z])([A-Z])", "$1 $2");
                    }

                    if (null != issue.Classification)
                        issue.ClassificationDisplayName = Regex.Replace(issue.Classification.ToString(), "([a-z])([A-Z])", "$1 $2");

                    if (null != issue.Origin)
                        issue.OriginDisplayName = Regex.Replace(issue.Origin.ToString(), "([a-z])([A-Z])", "$1 $2");

                    result.Add(issue);
                }

                if (idList.Count > 0)
                {
                    db.GetNotifications(idList, dbNotifications);
                    List<MSSummaryCacheEntry> summaryCache = db.GetLastNotificationSummary(idList);
                    foreach (MSServiceNotification dbNotification in dbNotifications)
                    {
                        M365ServiceIssue i = result?.Find(msg => msg.Id == dbNotification.Id);
                        if (i != null)
                            RehydratePersistedContent(i, dbNotification, summaryCache?.Find(sc => sc.id == dbNotification.Id));
                    }
                }
            }
            catch
            {

            }

            return result;
        }

        public List<M365ServiceIssue> GetServiceHealthIssueCollection()
        {
            List<M365ServiceIssue> serviceHealth;

            if (!m_cache.TryGetValue(c_serviceHealthCacheName, out serviceHealth))
            {
                serviceHealth = GetServiceHealthIssuesFromAPI();
                SetServiceHealthIssueCollection(serviceHealth);
            }

            return serviceHealth;
        }

        public void SetServiceHealthIssueCollection(List<M365ServiceIssue> value)
        {
            MemoryCacheEntryOptions options = new()
            {
                AbsoluteExpirationRelativeToNow = TimeSpan.FromMilliseconds(GlobalConfiguration.Instance.MemoryCacheExpiration),              
            };

            
            options.RegisterPostEvictionCallback((key, value, reason, substate) =>
            {
                var shhCollection = GetServiceHealthIssueCollection();
            });

            CancellationTokenSource token = new CancellationTokenSource();
            token.CancelAfter(GlobalConfiguration.Instance.MemoryCacheExpiration);

            options.AddExpirationToken(new CancellationChangeToken(token.Token));

            m_cache.Set(c_serviceHealthCacheName, value, options);
        }

        public M365ServiceIssue GetServiceHealthIssue(string id)
        {
            List<M365ServiceIssue> serviceHealth;
            M365ServiceIssue value;

            if (!m_cache.TryGetValue(c_serviceHealthCacheName, out serviceHealth))
            {
                serviceHealth = GetServiceHealthIssuesFromAPI();
                SetServiceHealthIssueCollection(serviceHealth);
            }
            
            value = serviceHealth?.Find(c => c.Id == id);

            return value;
        }
        public void SetServiceHealthIssue(M365ServiceIssue value)
        {
            List<M365ServiceIssue> serviceHealth;

            if (!m_cache.TryGetValue(c_serviceHealthCacheName, out serviceHealth))
            {
                serviceHealth = new();
                SetServiceHealthIssueCollection(serviceHealth);
            }

            M365ServiceIssue existingObject = serviceHealth?.Find(c => c.Id == value.Id);
            
            if (existingObject != null)
            {
                serviceHealth.Remove(existingObject);
            }

            serviceHealth.Add(value);
        }

        // region MessageCenter
        private void RehydratePersistedMessageCenterContent(M365MessageCenterItem message, MSServiceNotification persistedContent, MSSummaryCacheEntry summary)
        {
            try
            {
                message.Public = persistedContent.Public;
                message.PublishingComments = persistedContent.Comments;
                message.ExtendedProperties = persistedContent.ExtendedProperties;
                message.Summary = summary;
                message.ServiceHealthHubState = persistedContent.State;

                if (!string.IsNullOrWhiteSpace(persistedContent.WorkItemID))
                {
                    message.Task.TaskId = persistedContent.WorkItemID;
                    message.Task.TaskUrl = persistedContent.WorkItemURL;
                }
            }
            catch
            {

            }
        }

        private List<M365MessageCenterItem> GetMessageCenterItemsFromAPI()
        {
            List<ServiceUpdateMessage> allMessages = new List<ServiceUpdateMessage>();
            List<M365MessageCenterItem> result = new List<M365MessageCenterItem>();

            try
            {

                var apiCall = GraphApiClientHelper.Client.Admin.ServiceAnnouncement.Messages;

                Task<ServiceUpdateMessageCollectionResponse?> t = Task.Run(async () =>
                {
                    ServiceUpdateMessageCollectionResponse? res = await apiCall.GetAsync();
                    return res;
                });

                t.Wait();

                ServiceUpdateMessageCollectionResponse? messages = t.Result;

                allMessages.AddRange(messages?.Value);

                while (messages?.OdataNextLink != null)
                {
                    t = Task.Run(async () =>
                    {
                        ServiceUpdateMessageCollectionResponse? res = await apiCall.WithUrl(messages?.OdataNextLink).GetAsync();
                        return res;
                    });
                    t.Wait();
                    messages = t.Result;

                    allMessages.AddRange(messages?.Value);
                }

                List<string> idList = new List<string>();
                foreach (ServiceUpdateMessage m in allMessages)
                {
                    if (m.Id != null)
                        idList.Add(m.Id);
                }

                MSSHNotificationDatabase db = new();
                List<MSServiceNotification> dbNotifications = new List<MSServiceNotification>();

                foreach (var apiMessage in allMessages)
                {
                    M365MessageCenterItem message = M365MessageCenterItem.From(apiMessage);

                    if (message.AdditionalData == null)
                        message.AdditionalData = new Dictionary<string, object>();

                    result.Add(message);
                }

                if (idList.Count > 0)
                {
                    db.GetNotifications(idList, dbNotifications);
                    List<MSSummaryCacheEntry> summaryCache = db.GetLastNotificationSummary(idList);
                    foreach (MSServiceNotification dbNotification in dbNotifications)
                    {
                        M365MessageCenterItem m = result?.Find(msg => msg.Id == dbNotification.Id);
                        if (m != null)
                            RehydratePersistedMessageCenterContent(m, dbNotification, summaryCache?.Find(sc => sc.id == dbNotification.Id));
                    }
                }
            } catch
            {

            }

            return result;
        }

        public List<M365MessageCenterItem> GetMessageCenterCollection()
        {
            List<M365MessageCenterItem> messageCenter;

            if (!m_cache.TryGetValue(c_messageCenterCacheName, out messageCenter))
            {
                messageCenter = GetMessageCenterItemsFromAPI();
                SetMessageCenterCollection(messageCenter);
            }

            return messageCenter;
        }

        public void SetMessageCenterCollection(List<M365MessageCenterItem> value)
        {
            MemoryCacheEntryOptions options = new()
            {
                AbsoluteExpirationRelativeToNow = TimeSpan.FromMilliseconds(GlobalConfiguration.Instance.MemoryCacheExpiration),
            };


            options.RegisterPostEvictionCallback((key, value, reason, substate) =>
            {
                var shhCollection = GetMessageCenterCollection();
            });

            CancellationTokenSource token = new CancellationTokenSource();
            token.CancelAfter(GlobalConfiguration.Instance.MemoryCacheExpiration);

            options.AddExpirationToken(new CancellationChangeToken(token.Token));

            m_cache.Set(c_messageCenterCacheName, value, options);
        }

        public M365MessageCenterItem GetMessageCenterItem(string id)
        {
            List<M365MessageCenterItem> messageCenter;
            M365MessageCenterItem value;

            if (!m_cache.TryGetValue(c_messageCenterCacheName, out messageCenter))
            {
                messageCenter = GetMessageCenterItemsFromAPI();
                SetMessageCenterCollection(messageCenter);
            }

            value = messageCenter?.Find(c => c.Id == id);

            return value;
        }
        public void SetMessageCenterItem(M365MessageCenterItem value)
        {
            List<M365MessageCenterItem> messageCenter;

            if (!m_cache.TryGetValue(c_messageCenterCacheName, out messageCenter))
            {
                messageCenter = new();
                SetMessageCenterCollection(messageCenter);
            }

            M365MessageCenterItem existingObject = messageCenter?.Find(c => c.Id == value.Id);

            if (existingObject != null)
            {
                messageCenter.Remove(existingObject);
            }

            messageCenter.Add(value);
        }

        public T Get<T>(string key) where T : class
        {
            T value;

            if (!m_cache.TryGetValue(key, out value))
            {
                value = null;
            }

            return value;
        }

        public void Set<T>(string key, T value) where T : class
        {
            MemoryCacheEntryOptions options = new()
            {
                AbsoluteExpirationRelativeToNow = TimeSpan.FromMilliseconds(GlobalConfiguration.Instance.MemoryCacheExpiration)
            };

            m_cache.Set(key, value, options);
        }

        // Service Health Overview cache
        private List<ServiceHealth> GetServiceHealthOverviewFromAPI()
        {
            List<ServiceHealth> serviceStatus = new List<ServiceHealth>();

            try
            {
                var apiCall = GraphApiClientHelper.Client.Admin.ServiceAnnouncement.HealthOverviews;

                Task<ServiceHealthCollectionResponse?> t = Task.Run(async () =>
                {
                    ServiceHealthCollectionResponse? res = await apiCall.GetAsync();
                    return res;
                });

                t.Wait();

                ServiceHealthCollectionResponse? status = t.Result;

                serviceStatus.AddRange(status?.Value);

                while (status?.OdataNextLink != null)
                {
                    t = Task.Run(async () =>
                    {
                        ServiceHealthCollectionResponse? res = await apiCall.WithUrl(status?.OdataNextLink).GetAsync();
                        return res;
                    });
                    t.Wait();
                    status = t.Result;

                    serviceStatus.AddRange(status?.Value);
                }

                foreach (var apiStatus in serviceStatus)
                {
                    if (apiStatus.AdditionalData == null)
                        apiStatus.AdditionalData = new Dictionary<string, object>();

                    if (null != apiStatus.Status)
                        apiStatus.AdditionalData.Add("status", Regex.Replace(apiStatus.Status.ToString(), "([a-z])([A-Z])", "$1 $2"));
                }
            } 
            catch
            {

            }

            return serviceStatus;
        }

        public List<ServiceHealth> GetServiceHealthOverviewCollection()
        {
            List<ServiceHealth> serviceHealth;

            if (!m_cache.TryGetValue(c_serviceHealthOverviewCacheName, out serviceHealth))
            {
                serviceHealth = GetServiceHealthOverviewFromAPI();
                SetServiceHealthOverviewCollection(serviceHealth);
            }

            return serviceHealth;
        }

        public void SetServiceHealthOverviewCollection(List<ServiceHealth> value)
        {
            MemoryCacheEntryOptions options = new()
            {
                AbsoluteExpirationRelativeToNow = TimeSpan.FromMilliseconds(GlobalConfiguration.Instance.MemoryCacheExpiration),
            };


            options.RegisterPostEvictionCallback((key, value, reason, substate) =>
            {
                var shhCollection = GetServiceHealthIssueCollection();
            });

            CancellationTokenSource token = new();
            token.CancelAfter(GlobalConfiguration.Instance.MemoryCacheExpiration);

            options.AddExpirationToken(new CancellationChangeToken(token.Token));

            m_cache.Set(c_serviceHealthOverviewCacheName, value, options);
        }

        public static void Initialize()
        {
            Instance.GetServiceHealthIssueCollection();
            Instance.GetMessageCenterCollection();
            Instance.GetServiceHealthOverviewCollection();
        }
    }
}
