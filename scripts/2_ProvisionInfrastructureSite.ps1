# PROVISION INFRASTRUCTURE SITE
$InfrastructureSiteUrl = ""

# DO NOT MODIFY BELOW
$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning\OfficeDevPnP.PartnerPack.SiteProvisioning"

while($InfrastructureSiteUrl -eq "" -or $InfrastructureSiteUrl -eq $null -or $InfrastructureSiteUrl.ToLower() -notlike "https://*")
{
    $InfrastructureSiteUrl = Read-Host -Prompt "Enter infrastructure site url (e.g. https://yourtenant.sharepoint.com/sites/infrastructure)"
}

$creds = Get-Credential -Message "Enter Tenant Admin Credentials"

$uri = [System.Uri]$InfrastructureSiteUrl

$siteHost = $uri.Host.ToLower()
$siteHost = $siteHost.Replace(".sharepoint.com","-admin.sharepoint.com");

Connect-SPOnline -Url "https://$siteHost" -Credentials $creds
$infrastructureSiteInfo = Get-SPOTenantSite -Url $InfrastructureSiteUrl -ErrorAction SilentlyContinue
if($InfrastructureSiteInfo -eq $null)
{
    Write-Host -ForegroundColor Cyan "Infrastructure Site does not exist. Please create site collection first through the UI, or use New-SPOTenantSite"
} else {
    Connect-SPOnline -Url $InfrastructureSiteUrl -Credentials $creds
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
    
}