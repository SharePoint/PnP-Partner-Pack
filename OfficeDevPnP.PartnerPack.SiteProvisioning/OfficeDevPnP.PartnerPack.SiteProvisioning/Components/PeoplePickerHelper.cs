using Microsoft.SharePoint.ApplicationPages.ClientPickerQuery;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OfficeDevPnP.PartnerPack.Infrastructure;
using Microsoft.Graph;
using System.Threading.Tasks;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Models;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Components
{
    public class PeoplePickerHelper
    {
        private GraphServiceClient graphClient = MicrosoftGraphHelper.GetNewGraphClient();

        public async Task<PrincipalsViewModel> GetPeoplePickerSearchData(string nameToSearch, bool searchGroups = true, List<Guid> validationSkus = null, int maxResultCount = 6)
        {
            PrincipalsViewModel result = new PrincipalsViewModel();
            try
            {
                var filteredUsers = await GetUsers(nameToSearch);
                result.Principals.AddRange(filteredUsers.Principals);

                if (searchGroups)
                {
                    var filteredGroups = await GetGroups(nameToSearch);
                    result.Principals.AddRange(filteredGroups.Principals);
                }

                if (result.Principals.Count > maxResultCount)
                {
                    result.Principals = result.Principals.Take(maxResultCount).ToList();
                }
            }
            catch (Exception e)
            {
                // TODO: Handle exceptions with specific JSON response
                throw e;
            }

            return result;
        }

        private async Task<PrincipalsViewModel> GetUsers(string nameToSearch, IGraphServiceUsersCollectionRequest request = null)
        {
            PrincipalsViewModel result = new PrincipalsViewModel();

            var filteredUsers = await graphClient.Users.Request()
                    .Select("DisplayName,UserPrincipalName,Mail")
                    .Filter($"startswith(DisplayName,'{nameToSearch}') or startswith(UserPrincipalName,'{nameToSearch}') or startswith(Mail,'{nameToSearch}')")
                    .GetAsync();

            var t = await MapUsers(filteredUsers);
            result.Principals.AddRange(t.Principals);

            if (filteredUsers.NextPageRequest != null)
            {
                var additionalUsers = await GetUsers(nameToSearch, filteredUsers.NextPageRequest);

                if (additionalUsers.Principals.Count > 0)
                {
                    result.Principals.AddRange(additionalUsers.Principals);
                }
            }

            return result;
        }

        private async Task<PrincipalsViewModel> MapUsers(IGraphServiceUsersCollectionPage source)
        {
            PrincipalsViewModel result = new PrincipalsViewModel();

            if (source != null)
            {
                foreach (var u in source)
                {
                    u.Mail = (string.IsNullOrEmpty(u.Mail)) ? u.UserPrincipalName : u.Mail;
                    result.Principals.Add(await MapUser(u));
                }
            }

            return result;
        }

        private async Task<PrincipalViewModel> MapUser(Microsoft.Graph.User u)
        {
            if (u == null)
            {
                return null;
            }

            PrincipalViewModel user = new PrincipalViewModel()
            {
                UserPrincipalName = u.UserPrincipalName,
                DisplayName = u.DisplayName,
                Mail = u.Mail,
                FirstName = u.GivenName,
                LastName = u.Surname,
                JobTitle = u.JobTitle
            };

            return user;
        }

        private async Task<PrincipalsViewModel> GetGroups(string nameToSearch, IGraphServiceGroupsCollectionRequest request = null)
        {
            PrincipalsViewModel result = new PrincipalsViewModel();

            var filteredGroups = await graphClient.Groups.Request()
                    .Select("DisplayName,Description")
                    .Filter("groupTypes/any(grp: grp ne 'Unified')") //DynamicMembership 
                    .GetAsync();

            var t = await MapGroups(filteredGroups);
            result.Principals.AddRange(t.Principals);

            if (filteredGroups.NextPageRequest != null)
            {
                var additionalGroups = await GetGroups(nameToSearch, filteredGroups.NextPageRequest);

                if (additionalGroups.Principals.Count > 0)
                {
                    result.Principals.AddRange(additionalGroups.Principals);
                }
            }

            return result;
        }

        private async Task<PrincipalsViewModel> MapGroups(IGraphServiceGroupsCollectionPage source)
        {
            PrincipalsViewModel result = new PrincipalsViewModel();

            if (source != null)
            {
                foreach (var g in source)
                {
                    result.Principals.Add(await MapGroup(g));
                }
            }

            return result;
        }

        private async Task<PrincipalViewModel> MapGroup(Microsoft.Graph.Group g)
        {
            if (g == null)
            {
                return null;
            }

            PrincipalViewModel group = new PrincipalViewModel()
            {
                DisplayName = g.DisplayName,
                Mail = g.Mail,
                JobTitle = "Group"
            };

            return group;
        }
    }
}