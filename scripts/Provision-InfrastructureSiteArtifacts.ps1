<#
.SYNOPSIS
Provisions required artifacts to the Infrastructure Site

.EXAMPLE
PS C:\> .\Provision-InfrastructureSiteArtifacts.ps1 -InfastructureSiteUrl "https://mytenant.sharepoint.com/sites/infrastructure"

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\>.\Provision-InfrastructureSiteArtifacts.ps1 -InfastructureSiteUrl "https://mytenant.sharepoint.com/sites/infrastructure" -Credentials $creds

#>
[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage="The URL of your infrastructure site, e.g. https://mytenant.sharepoint.com/sites/infrastructure")]
    [String]
    $InfrastructureSiteUrl,

	[Parameter(Mandatory = $false, HelpMessage="Optional tenant administration credentials")]
	[PSCredential]
	$Credentials

)


# DO NOT MODIFY BELOW
$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning\OfficeDevPnP.PartnerPack.SiteProvisioning"

if($Credentials -eq $null)
{
	$Credentials = Get-Credential -Message "Enter Tenant Admin Credentials"
}

$uri = [System.Uri]$InfrastructureSiteUrl

$siteHost = $uri.Host.ToLower()
$siteHost = $siteHost.Replace(".sharepoint.com","-admin.sharepoint.com");

Connect-SPOnline -Url "https://$siteHost" -Credentials $Credentials
$infrastructureSiteInfo = Get-SPOTenantSite -Url $InfrastructureSiteUrl -ErrorAction SilentlyContinue
if($InfrastructureSiteInfo -eq $null)
{
    Write-Host -ForegroundColor Cyan "Infrastructure Site does not exist. Please create site collection first through the UI, or use New-SPOTenantSite"
} else {
    Connect-SPOnline -Url $InfrastructureSiteUrl -Credentials $Credentials
    Apply-SPOProvisioningTemplate -Path "$basePath\Templates\Infrastructure\PnP-Partner-Pack-Infrastructure-Jobs.xml"
    Apply-SPOProvisioningTemplate -Path "$basePath\Templates\Infrastructure\PnP-Partner-Pack-Infrastructure-Templates.xml"

    # Unhide the 2 infrastructure lists
    $l = Get-SPOList "PnPProvisioningTemplates"
    $l.Hidden = $false
    $l.OnQuickLaunch = $true
    $l.Update()
    Execute-SPOQuery

    $l = Get-SPOList "PnPPRovisioningJobs"
    $l.Hidden = $false
    $l.OnQuickLaunch = $true
    $l.Update()
    Execute-SPOQuery
    
	Apply-SPOProvisioningTemplate $basePath\Templates\PnP-Partner-Pack-Infrastructure-Contents.xml
}