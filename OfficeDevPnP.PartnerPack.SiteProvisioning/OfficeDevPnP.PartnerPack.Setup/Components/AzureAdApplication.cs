using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class KeyCredential
    {
        public string customKeyIdentifier { get; set; }
        public string keyId { get; set; }
        public string type { get; set; }
        public string usage { get; set; }
        public string value { get; set; }
    }

    public class ResourceAccess
    {
        public string id { get; set; }
        public string type { get; set; }
    }

    public class RequiredResourceAccess
    {
        public string resourceAppId { get; set; }
        public List<ResourceAccess> resourceAccess { get; set; }
    }

    public class AzureAdApplication
    {
        public Guid? Id { get; set; }
        public List<object> addIns { get; set; }
        public List<object> appRoles { get; set; }
        public Guid? AppId { get; set; }
        public bool availableToOtherOrganizations { get; set; }
        public string displayName { get; set; }
        public object errorUrl { get; set; }
        public object groupMembershipClaims { get; set; }
        public string homepage { get; set; }
        public List<string> identifierUris { get; set; }
        public List<KeyCredential> keyCredentials { get; set; }
        public List<object> knownClientApplications { get; set; }
        public object logoutUrl { get; set; }
        public bool oauth2AllowImplicitFlow { get; set; }
        public bool oauth2AllowUrlPathMatching { get; set; }
        public List<object> oauth2Permissions { get; set; }
        public bool oauth2RequirePostResponse { get; set; }
        public List<object> passwordCredentials { get; set; }
        public bool publicClient { get; set; }
        public List<string> replyUrls { get; set; }
        public List<RequiredResourceAccess> requiredResourceAccess { get; set; }
        public object samlMetadataUrl { get; set; }
    }
}
