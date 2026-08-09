namespace Microsoft.ServiceHealthHub.Core
{
    public class CommunicationBase
    {
        public string WorkItemId { get; set; }
        public string WorkItemUrl { get; set; }
        public dynamic Message { get; set; }

        public CommunicationBase()
        {

        }

        public CommunicationBase(string id)
        {

        }
    }
}
