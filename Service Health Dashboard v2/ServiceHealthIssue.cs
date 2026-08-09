using System;
using System.Collections.Generic;

namespace Microsoft.Util.GraphAPI
{
	public class ServiceHealthIssuePostItemBody
    {
		public string contentType { get; set; }
		public string content { get; set; }
	}
	
	public class ServiceHealthIssuePost
    {
		public DateTimeOffset createdDateTime { get; set; }
		public ServiceHealthIssuePostItemBody description { get; set; }
	}

	public class ServiceHealthIssue
	{
		public string id { get; set; }
		public string titme { get; set; }
		public string status { get; set; }
		public bool isResolved { get; set; }
		public string classification { get; set; }
		public string origin { get; set; }
		public string service { get; set; }
		public string feature { get; set; }
		public string featureGroup { get; set; }
		public string impactDescription { get; set; }
		public IEnumerable<IDictionary<string, object>> details { get; set; }	
		public List<ServiceHealthIssuePost> posts { get; set; }
		public DateTimeOffset? startDateTime { get; set; }
		public DateTimeOffset? endDateTime { get; set; }
		public DateTimeOffset? lastModifiedDateTime { get; set; }
		public string SHD_WorkItemID { get; set; }
		public string SHD_WorkItemURL { get; set; }
	}

	public class ServiceHealthIssueCollection
	{
		public List<ServiceHealthIssue> value { get; set; }
	}
}
