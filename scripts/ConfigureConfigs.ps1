# The name of your storage account
$AzureStorageAccountName = ""; 
# The primary account key of your Azure Storage instance
$AzureStoragePrimaryAccountKey = "";
# The Client Id
$ClientId= ""
# The Client Secret
$ClientSecret = ""
# Your AD Tenant
$ADTenant = ""
# The full path to the certificate file that you uploaded to the azure web site for app authentication
$CertificatePath = ""
# Or the certificate thumbprint value
$CertificateThumbprint = ""
# The full url to your tenant administration site
$InfrastructureSiteUrl = ""

# DO NOT MODIFY BELOW
$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning"
$configFiles =  "OfficeDevPnP.PartnerPack.CheckAdminsJob\App.config",
                "OfficeDevPnP.PartnerPack.ExternalUsersJob\App.config",
                "OfficeDevPnP.PartnerPack.ContinousJob\App.config",
                "OfficeDevPnP.PartnerPack.ScheduledJob\App.config",
                "OfficeDevPnP.PartnerPack.SiteProvisioning\Web.config"

$azureWebJobsDashBoardNodeValue = "DefaultEndpointsProtocol=https;AccountName=$AzureStorageAccountName;AccountKey=$AzureStoragePrimaryAccountKey";

foreach($configFile in $configFiles)
{   
    $configDoc = (Get-Content "$basePath\$configFile") -As [Xml]
    $azureWebJobsDashboardNode = $configDoc.configuration.connectionStrings.add | ?{$_.name -eq "AzureWebJobsDashboard"}
    if($azureWebJobsDashboardNode -ne $null)
    {
        $azureWebJobsDashboardNode.connectionString = $azureWebJobsDashBoardNodeValue
    }
    $azureWebJobsStorageNode = $configDoc.configuration.connectionStrings.add | ?{$_.name -eq "AzureWebJobsStorage"}
    if($azureWebJobsStorageNode -ne $null)
    {
        $azureWebJobsStorageNode.connectionString = $azureWebJobsDashBoardNodeValue
    }
    $clientIdNode = $configDoc.configuration.appSettings.add | ? {$_.key -eq "ida:ClientId"}
    $clientIdNode.value = $ClientId
    $clientSecretNode = $configDoc.configuration.appSettings.add | ? {$_.key -eq "ida:ClientSecret"}
    $clientSecretNode.value = $ClientSecret

    $tenantSettingsNode = $configDoc.configuration.PnPPartnerPackConfiguration.TenantSettings
    $tenantSettingsNode.tenant = $ADTenant
    if ($CertificateThumbprint -ne "") 
    {
        $tenantSettingsNode.appOnlyCertificateThumbprint = $CertificateThumbprint
    }
    else 
    {
        $cert = Get-PfxCertificate $CertificatePath
        $tenantSettingsNode.appOnlyCertificateThumbprint = $cert.Thumbprint
    }
    $tenantSettingsNode.infrastructureSiteUrl = $InfrastructureSiteUrl

    $configDoc.Save("$basePath\$configFile")
}