using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class RequiredIfAttribute : RequiredAttribute
    {
        public string DependantPropertyName { get; set; }

        public RequiredIfAttribute(string dependantPropertyName)
        {
            if (dependantPropertyName == null)
                throw new ArgumentNullException(nameof(dependantPropertyName));
            DependantPropertyName = dependantPropertyName;
        }

        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            var property = TypeDescriptor.GetProperties(validationContext.ObjectInstance).Find(DependantPropertyName, true);
            if (property == null)
                throw new InvalidOperationException("Cannot find property " + DependantPropertyName);

            var dependantValue = property.GetValue(validationContext.ObjectInstance);
            if (Convert.ToBoolean(dependantValue))
            {
                return base.IsValid(value, validationContext);
            }

            return ValidationResult.Success;
        }
    }
}
