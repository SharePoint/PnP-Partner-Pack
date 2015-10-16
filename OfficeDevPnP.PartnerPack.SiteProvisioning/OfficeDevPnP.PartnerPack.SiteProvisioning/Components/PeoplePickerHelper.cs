using Microsoft.SharePoint.ApplicationPages.ClientPickerQuery;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OfficeDevPnP.PartnerPack.Infrastructure;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Components
{
    public class PeoplePickerHelper
    {
        private static int GroupID = -1;

        public static string GetPeoplePickerSearchData()
        {
            using (var context = PnPPartnerPackContextProvider.GetAppOnlyClientContext(PnPPartnerPackSettings.InfrastructureSiteUrl))
            {
                return GetPeoplePickerSearchData(context);
            }
        }

        public static string GetPeoplePickerSearchData(ClientContext context)
        {
            //get searchstring and other variables
            var searchString = (string)HttpContext.Current.Request["SearchString"];
            int principalType = Convert.ToInt32(HttpContext.Current.Request["PrincipalType"]);
            string spGroupName = (string)HttpContext.Current.Request["SPGroupName"];

            ClientPeoplePickerQueryParameters querryParams = new ClientPeoplePickerQueryParameters();
            querryParams.AllowMultipleEntities = false;
            querryParams.MaximumEntitySuggestions = 2000;
            querryParams.PrincipalSource = PrincipalSource.All;
            querryParams.PrincipalType = (PrincipalType)principalType;
            querryParams.QueryString = searchString;

            if (!string.IsNullOrEmpty(spGroupName))
            {
                if (PeoplePickerHelper.GroupID == -1)
                {
                    var group = context.Web.SiteGroups.GetByName(spGroupName);
                    if (group != null)
                    {
                        context.Load(group, p => p.Id);
                        context.ExecuteQuery();

                        PeoplePickerHelper.GroupID = group.Id;

                        querryParams.SharePointGroupID = group.Id;
                    }
                }
                else
                {
                    querryParams.SharePointGroupID = PeoplePickerHelper.GroupID;
                }
            }

            //execute query to Sharepoint
            ClientResult<string> clientResult = Microsoft.SharePoint.ApplicationPages.ClientPickerQuery.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser(context, querryParams);
            context.ExecuteQuery();
            return clientResult.Value;
        }
    }
}