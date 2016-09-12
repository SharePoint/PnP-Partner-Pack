using Newtonsoft.Json;
using OfficeDevPnP.PartnerPack.SiteProvisioning.Binders;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace OfficeDevPnP.PartnerPack.SiteProvisioning.Models
{
    /// <summary>
    /// Class used to identity a User
    /// </summary>
    [ModelBinder(typeof(PrincipalViewModelBinder))]
    public class PrincipalViewModel
    {
        public PrincipalViewModel()
        {
            string[] colors = new string[]
            {
                    "black",
                    "blue",
                    "darkBlue",
                    "darkGreen",
                    "darkRed",
                    "green",
                    "lightBlue",
                    "lightGreen",
                    "lightPink",
                    "magenta",
                    "orange",
                    "pink",
                    "purple",
                    "red",
                    "teal"
            };

            Random r = new Random(DateTime.Now.Millisecond);
            int index = r.Next(0, colors.Count() - 1);

            String prefix = "ms-Persona-initials--";

            String result = $"{prefix}{colors[index]}";

            BadgeColor = result;
        }

        /// <summary>
        /// Email
        /// </summary>
        [JsonProperty(PropertyName = "mail")]
        public String Mail { get; set; }

        /// <summary>
        /// User Principal Name
        /// </summary>
        [JsonProperty(PropertyName = "userPrincipalName")]
        public String UserPrincipalName { get; set; }

        /// <summary>
        /// Display name
        /// </summary>
        [JsonProperty(PropertyName = "displayName")]
        public String DisplayName { get; set; }

        /// <summary>
        /// Job title
        /// </summary>
        [JsonProperty(PropertyName = "jobTitle")]
        public String JobTitle { get; set; }

        /// <summary>
        /// First name
        /// </summary>
        [JsonProperty(PropertyName = "givenName")]
        public String FirstName { get; set; }

        /// <summary>
        /// Last name
        /// </summary>
        [JsonProperty(PropertyName = "surname")]
        public String LastName { get; set; }

        private String _abbreviation;

        public String Abbreviation
        {
            get
            {
                if (String.IsNullOrEmpty(_abbreviation))
                {
                    if (!String.IsNullOrEmpty(FirstName) && !String.IsNullOrEmpty(LastName))
                    {
                        _abbreviation = $"{FirstName.First()}{LastName.First()}".ToUpper();
                    }
                    else if (!String.IsNullOrEmpty(DisplayName))
                    {
                        var splitted = DisplayName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        // length > 2 to avoid strings like 'e', '-', 'to'
                        _abbreviation = String.Concat(splitted.Where(a => a.Length > 2).Select(a => a.First())).ToUpper();
                    }
                }

                return _abbreviation;
            }
        }

        public String BadgeColor { get; set; }
    }

    public class PrincipalsViewModel : IValidatableObject
    {
        /// <summary>
        /// Represents the selected principals
        /// </summary>
        public List<PrincipalViewModel> Principals { get; set; } = new List<PrincipalViewModel>();

        /// <summary>
        /// Defines the maximum number of selectable profiles
        /// </summary>
        public Int32 MaxSelectableProfilesNumber { get; set; } = 1;

        /// <summary>
        /// Allows to defines whether to search for groups, and not only for users
        /// </summary>
        public Boolean IncludeGroups { get; set; } = false;

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if (Principals == null || Principals.Count == 0)
                yield return new ValidationResult("Please select at least one principal");
        }
    }
}