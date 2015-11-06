#################################################################################################
#
#
# Use this script to reset the app.config and web.config files to their original unmodified state
#
#
#################################################################################################

$ClientId= "[CLIENT ID]"
$ClientSecret = "[CLIENT SECRET]"
$ADTenant = "[TENANT].onmicrosoft.com"
$InfrastructureSiteUrl = "[INFRASTRUCTURE SITE URL]"

# DO NOT MODIFY BELOW
$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning"
$configFiles =  "OfficeDevPnP.PartnerPack.CheckAdminsJob\App.config",
                "OfficeDevPnP.PartnerPack.ExternalUsersJob\App.config",
                "OfficeDevPnP.PartnerPack.ContinousJob\App.config",
                "OfficeDevPnP.PartnerPack.ScheduledJob\App.config",
                "OfficeDevPnP.PartnerPack.SiteProvisioning\Web.config"

$azureWebJobsDashBoardNodeValue = "";

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
    $tenantSettingsNode.appOnlyCertificateThumbprint = "[CERTIFICATE THUMBPRINT]"
    $tenantSettingsNode.infrastructureSiteUrl = $InfrastructureSiteUrl

    $configDoc.Save("$basePath\$configFile")
}