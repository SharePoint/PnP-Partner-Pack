﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    /// <summary>
    /// Defines the setup settings
    /// </summary>
    public class SetupInformation
    {
        public String ApplicationName { get; set; }
        public String ApplicationUniqueUri { get; set; }
        public String ApplicationLogoPath { get; set; }
        public String AzureWebAppUrl { get; set; }
        public Boolean SslCertificateGenerate { get; set; }
        public String SslCertificateFile { get; set; }
        public String SslCertificatePassword { get; set; }
        public String SslCertificateCommonName { get; set; }
        public DateTime SslCertificateStartDate { get; set; }
        public DateTime SslCertificateEndDate { get; set; }
        public String InfrastructuralSiteUrl { get; set; }
        public Int32 InfrastructuralSiteLCID { get; set; }
        public Int32 InfrastructuralSiteTimeZone { get; set; }
        public Guid AzureTargetSubscription { get; set; }
        public String AzureAppServiceName { get; set; }
        public String AzureBlobStorageName { get; set; }
    }
}