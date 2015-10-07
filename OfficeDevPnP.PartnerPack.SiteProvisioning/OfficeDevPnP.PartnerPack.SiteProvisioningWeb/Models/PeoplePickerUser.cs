using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeDevPnP.PartnerPack.SiteProvisioningWeb.Models
{
    /// <summary>
    /// This type is used for collection users from PeoplePicker client-side control
    /// </summary>
    public class PeoplePickerUser
    {
        /// <summary>
        /// User's login name
        /// </summary>
        public String Login { get; set; }

        /// <summary>
        /// User's name
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// User's email
        /// </summary>
        public String Email { get; set; }
    }
}