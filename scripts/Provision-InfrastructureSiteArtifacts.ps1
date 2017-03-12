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

    [Parameter(Mandatory = $true, HelpMessage="The URL of your Azure Web site")]
    [String]
    $AzureWebSiteUrl,

	[Parameter(Mandatory = $false, HelpMessage="Optional tenant administration credentials")]
	[PSCredential]
	$Credentials

)


# DO NOT MODIFY BELOW
$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning\OfficeDevPnP.PartnerPack.SiteProvisioning"

# Modify Responsive design template to include Azure WebSite Url
$responsiveTemplate = (Get-Content "$basePath\Templates\Responsive\SPO-Responsive.xml") -As [Xml]
$parameter = $responsiveTemplate.Provisioning.Preferences.Parameters.Parameter | ?{$_.Key -eq "AzureWebSiteUrl"}
$parameter.InnerText = $AzureWebSiteUrl.Trim('/')
$responsiveTemplate.Save("$basePath\Templates\Responsive\SPO-Responsive.xml");

# Modify PnP Partner Pack Overrides template to include Azure WebSite Url
$responsiveTemplate = (Get-Content "$basePath\Templates\Overrides\PnP-Partner-Pack-Overrides.xml") -As [Xml]
$parameter = $responsiveTemplate.Provisioning.Preferences.Parameters.Parameter | ?{$_.Key -eq "AzureWebSiteUrl"}
$parameter.InnerText = $AzureWebSiteUrl.Trim('/')
$responsiveTemplate.Save("$basePath\Templates\Overrides\PnP-Partner-Pack-Overrides.xml");

if($Credentials -eq $null)
{
	$Credentials = Get-Credential -Message "Enter Tenant Admin Credentials"
}

$InfrastructureSiteUrl = $InfrastructureSiteUrl.Trim('/')
$uri = [System.Uri]$InfrastructureSiteUrl

$siteHost = $uri.Host.ToLower()
$siteHost = $siteHost.Replace(".sharepoint.com","-admin.sharepoint.com")
$siteHost = $siteHost.Trim('/')

Connect-SPOnline -Url "https://$siteHost" -Credentials $Credentials
$infrastructureSiteInfo = Get-PnPTenantSite -Url $InfrastructureSiteUrl -ErrorAction SilentlyContinue
if($InfrastructureSiteInfo -eq $null)
{
    Write-Host -ForegroundColor Cyan "Infrastructure Site does not exist. Please create site collection first through the UI, or use New-PnPTenantSite"
} else {
    Connect-SPOnline -Url $InfrastructureSiteUrl -Credentials $Credentials
    Apply-PnPProvisioningTemplate -Path "$basePath\Templates\Infrastructure\PnP-Partner-Pack-Infrastructure-Jobs.xml"
    Apply-PnPProvisioningTemplate -Path "$basePath\Templates\Infrastructure\PnP-Partner-Pack-Infrastructure-Templates.xml"

    # Unhide the 2 infrastructure lists
    $l = Get-PnPList "PnPProvisioningTemplates"
    $l.Hidden = $false
    $l.OnQuickLaunch = $true
    $l.Update()
    Execute-PnPQuery

    $l = Get-PnPList "PnPPRovisioningJobs"
    $l.Hidden = $false
    $l.OnQuickLaunch = $true
    $l.Update()
    Execute-PnPQuery
    
    Apply-PnPProvisioningTemplate $basePath\Templates\PnP-Partner-Pack-Infrastructure-Contents.xml

	
    # Due to an issue in core, we have to loop through the items in the provisioning template list and set the correct content type
    # Find the content type in the list
    $ctx = Get-PnPContext
    $l = Get-PnPList "PnPProvisioningTemplates"
    $ctx.Load($l.ContentTypes)
    Execute-PnPQuery
    $ct = $l.ContentTypes | ?{$_.Name -eq "PnpProvisioningTemplate"}

    $items = Get-PnPListItem -List "PnPProvisioningTemplates" -Fields "FileLeafRef"
    foreach($item in $items)
    {
	$filename = $item["FileLeafRef"]
	if($filename -eq "SPO-CustomBar.xml" -or $filename -eq "PnP-Partner-Pack-Overrides.xml" -or $filename -eq "SPO-Responsive.xml" -or $filename -eq "PnP-Partner-Pack-Infrastructure-Jobs.xml" -or $filename -eq "PnP-Empty-Site.xml" -or $filename -eq "PnP-Empty-Sub-Site.xml")
	{
		Write-Host -ForeGroundColor DarkGray "Setting content type on $filename"
		$item["ContentTypeId"] = $ct.Id
		$item.Update()
		Execute-PnPQuery
	}
    }
    
}