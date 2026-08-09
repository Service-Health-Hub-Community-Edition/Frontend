using Microsoft.Data.SqlClient.Server;
using System.Data;

namespace Microsoft.ServiceHealthHub.Core
{
    public class MSCommunicationId
    {
        public Guid userId { get; set; }
        public string communicationId { get; set; }
        public Guid tenantId { get; set; }
        public Guid subscriptionId { get; set; }
    }

    public class MSCommunicationIdCollection: List<MSCommunicationId>, IEnumerable<SqlDataRecord>
    {
        IEnumerator<SqlDataRecord> IEnumerable<SqlDataRecord>.GetEnumerator()
        {
            SqlDataRecord ret = new SqlDataRecord(
                new SqlMetaData("UserId", SqlDbType.UniqueIdentifier),
                new SqlMetaData("CommunicationId", SqlDbType.NVarChar, 64),
                new SqlMetaData("TenantId", SqlDbType.UniqueIdentifier),
                new SqlMetaData("SubscriptionId", SqlDbType.UniqueIdentifier)
                );

            foreach (MSCommunicationId item in this )
            {
                ret.SetSqlGuid(0, item.userId);
                ret.SetString(1, item.communicationId);
                ret.SetSqlGuid(2, item.tenantId);
                ret.SetSqlGuid(3, item.subscriptionId);
                yield return ret;
            }
        }
    }

    public class MSViewpoint
    {
        public Guid? userId { get; set; }
        public string communicationId { get; set; }
        public Guid? tenantId { get; set; }
        public Guid? subscriptionId { get; set; }
        public bool? read { get; set; }
        public bool? favorite { get; set; }
        public bool? archived { get; set; }

        public MSViewpoint() { }

        public MSViewpoint(Guid userId, string communicationId, Guid tenantId, Guid subscriptionId)
        {
        
        }
    }
}
