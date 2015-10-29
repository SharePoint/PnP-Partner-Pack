# The name of your storage account
$AzureStorageAccountName = "erwinvanhunen"; 
# The primary account key of your Azure Storage instance
$AzureStoragePrimaryAccountKey = "m3ePix42HPUfs2Ul8zksiwNm/7lYOXFrj2wuhcZ4QM40ELFxhhTn3G/NaNKtdL8o1NYLeICvywER7sycMCKg9g==";
# The Client Id
$ClientId= "c4d6366f-8625-4f7a-b745-bd0f3597a782"
# The Client Secret
$ClientSecret = "1idhGLO64QSNeaXVcP3JyddMZ9VDixl75cppgdQwbgE="
# Your AD Tenant
$ADTenant = "erwinmcm.onmicrosoft.com"
# The full path to the certificate file that you uploaded to the azure web site for app authentication
$CertificatePath = "c:\temp\pnp-partnerpack.cer"
# The full url to your tenant administration site
$InfrastructureSiteUrl = "https://erwinmcm.sharepoint.com/sites/pnp-partner-pack"

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
    $cert = Get-PfxCertificate $CertificatePath
    $tenantSettingsNode.appOnlyCertificateThumbprint = $cert.Thumbprint
    $tenantSettingsNode.infrastructureSiteUrl = $InfrastructureSiteUrl

    $configDoc.Save("$basePath\$configFile")
}