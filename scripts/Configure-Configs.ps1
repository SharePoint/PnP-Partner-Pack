<#
.SYNOPSIS
Configures the projects in the PnP PartnerPack Solution to use the correct values in app.config and web.config files

.EXAMPLE
PS C:\> .\Configure-Configs.ps1 `
-AzureStorageAccountName "MyStorageAccount" `
-AzureStoragePrimaryAccessKey  "m3ePix4-------------------------== `
-ClientId "0ce12c51-ac0c-4dba-9491-ca78b17dd8a0" `
-ClientSecret "aIkj5T6PYBa-------=" `
-ADTenant "mytenant.onmicrosoft.com" `
-CertificatePath "c:\temp\mycertificate.pfx" `
-InfastructureSiteUrl "https://mytenant.sharepoint.com/sites/infrastructure"

.EXAMPLE
PS C:\> .\Configure-Configs.ps1 `
-AzureStorageAccountName "MyStorageAccount" `
-AzureStoragePrimaryAccessKey  "m3ePix4-------------------------== `
-ClientId "0ce12c51-ac0c-4dba-9491-ca78b17dd8a0" `
-ClientSecret "aIkj5T6PYBa-------=" `
-ADTenant "mytenant.onmicrosoft.com" `
-CertificateThumbprint "B8D26235D65A4DE6ABF3D0DFC31C144D" `
-InfastructureSiteUrl "https://mytenant.sharepoint.com/sites/infrastructure"
#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the name of your storage account, e.g. 'mystorageaccount'")]
    [String]
    $AzureStorageAccountName,

    [Parameter(Mandatory = $true, HelpMessage="Enter the primare account key for your storage account")]
    [String]
    $AzureStoragePrimaryAccessKey,

	[Parameter(Mandatory = $true, HelpMessage="Enter the client id for your Azure AD App")]
    [String]
    $ClientId,

    [Parameter(Mandatory = $true, HelpMessage="Enter the client secret for your Azure AD App")]
    [String]
    $ClientSecret,

    [Parameter(Mandatory = $true, HelpMessage="Enter the AD tenant name, e.g. mytenant.onmicrosoft.com")]
	[String]
	$ADTenant,

	[Parameter(Mandatory = $false, HelpMessage="The full path to the certificate file that you uploaded to the azure web site for app authentication", ParameterSetName="CertificatePath")]
    [String]
    $CertificatePath,

	[Parameter(Mandatory = $false, HelpMessage="The certificate thumbprint value", ParameterSetName="Thumbprint")]
    [String]
    $CertificateThumbprint,

	[Parameter(Mandatory = $true)]
    [String]
    $InfrastructureSiteUrl

)
# DO NOT MODIFY BELOW
$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning"
$configFiles =  "OfficeDevPnP.PartnerPack.CheckAdminsJob\App.config",
                "OfficeDevPnP.PartnerPack.ExternalUsersJob\App.config",
                "OfficeDevPnP.PartnerPack.ContinousJob\App.config",
                "OfficeDevPnP.PartnerPack.ScheduledJob\App.config",
                "OfficeDevPnP.PartnerPack.SiteProvisioning\Web.config"


$azureWebJobsDashBoardNodeValue = "DefaultEndpointsProtocol=https;AccountName=$AzureStorageAccountName;AccountKey=$AzureStoragePrimaryAccessKey";

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