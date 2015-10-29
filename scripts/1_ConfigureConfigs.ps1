# The name of your storage account
$AzureStorageAccountName = ""; 
# The primary account key of your Azure Storage instance
$AzureStoragePrimaryAccountKey = "";
# The Client Id
$ClientId= ""
# The Client Secret
$ClientSecret = ""
# Your AD Tenant, e.g. "mytenant.onmicrosoft.com"
$ADTenant = ""
# The full path to the certificate file that you uploaded to the azure web site for app authentication
$CertificatePath = ""
# Or the certificate thumbprint value
$CertificateThumbprint = ""
# The full url to your infrastructure site
$InfrastructureSiteUrl = ""

# DO NOT MODIFY BELOW
$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning"
$configFiles =  "OfficeDevPnP.PartnerPack.CheckAdminsJob\App.config",
                "OfficeDevPnP.PartnerPack.ExternalUsersJob\App.config",
                "OfficeDevPnP.PartnerPack.ContinousJob\App.config",
                "OfficeDevPnP.PartnerPack.ScheduledJob\App.config",
                "OfficeDevPnP.PartnerPack.SiteProvisioning\Web.config"

while($AzureStorageAccountName -eq "" -or $AzureStorageAccountName -eq $null)
{
    $AzureStorageAccountName = Read-Host -Prompt "Enter name of Storage Account (e.g. mystorageaccount)"
}

while($AzureStoragePrimaryAccountKey -eq "" -or $AzureStoragePrimaryAccountKey -eq $null)
{
    $AzureStoragePrimaryAccountKey = Read-Host -Prompt "Enter primary account key of $AzureStorageAccount"
}

while($ClientId -eq "" -or $ClientId -eq $null)
{
    $ClientId = Read-Host -Prompt "Enter Client Id"
}

while($ClientSecret -eq "" -or $ClientSecret -eq $null)
{
    $ClientSecret = Read-Host -Prompt "Enter Client Secret"
}

while($ADTenant -eq "" -or $ADTenant -eq $null)
{
    $ADTenant = Read-Host -Prompt "Enter AD Tenant (e.g. mytenant.onmicrosoft.com)"
}

if($CertificatePath -eq "" -or $CertificatePath -eq $null)
{
    if($CertificateThumbprint -eq "" -or $CertificateThumbprint -eq $null)
    {
        $CertificatePath = Read-Host "Enter full path to certificate. Leave empty to entry certificate thumbprint instead"
        if($CertificatePath -eq "")
        {
            while($CertificateThumbprint -eq "" -or $CertificateThumbprint -eq $null)
            {
                $CertificateThumbprint = Read-Host "Enter certificate thumbprint"
            }
        }    
    }
}

while($InfrastructureSiteUrl -eq "" -or $InfrastructureSiteUrl -eq $null)
{
    $InfrastructureSiteUrl = Read-Host -Prompt "Enter full url to your infrastructure site (e.g. https://mytenant.sharepoint.com/sites/infrastructure)"
}

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