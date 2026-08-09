using System.Security.Claims;

namespace Microsoft.ServiceHealthHub.Core
{
    public class MSUserInfo
    {
        private string fullName = string.Empty;
        private string userName = string.Empty;
        private string objectId = string.Empty;

        public string FullName => fullName;
        public string UserName => userName;
        public string ObjectId => objectId;

        public MSUserInfo(ClaimsPrincipal User)
        {
            List<Claim> claims = User.Identities.First()?.Claims.ToList();
            fullName = claims.Find(c => c.Type == "name")?.Value;
            userName = claims.Find(c => c.Type == "preferred_username")?.Value;
            objectId = claims.Find(c => c.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value;
        }
    }
}
