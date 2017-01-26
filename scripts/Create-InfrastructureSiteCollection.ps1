[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the name of your tenant, e.g. 'contoso'")]
    [String]
    $Tenant, 
     [Parameter(Mandatory = $true, HelpMessage="Enter the Full url of the Infrastructure site , e.g. 'https://contoso.sharepoint.com/sites/PnP-Partner-Pack-Infrastructure'")]
    [String]
    $InfrastructureSiteUrl,     
    [Parameter(Mandatory = $true, HelpMessage="Enter the name of site collection owner, e.g. 'admin@contoso.com'")]
    [String]
    $Owner,
    [Parameter(Mandatory = $true, HelpMessage="Enter the name of azure app service , e.g. 'PnPPartnerPack'")]
    [String]
    $AzureService,
    [Parameter(Mandatory = $false, HelpMessage="Office 365 Creds ")]
    $Credentials  
)

Write-Host "Creating Infrastructure Site collection. It will wait until it is finished"
$job  = Start-Job { 
    Connect-SPOnline "https://$tenant-admin.sharepoint.com/" -Credentials $Credentials
    if((Get-SPOTenantSite -Url $InfrastructureUrl -ErrorAction 0) -eq $null)
    {
        New-SPOTenantSite -Title "PnP Partner Pack - Infrastructural Site" -Url $InfrastructureSiteUrl -Owner $Owner -Lcid 1033 -Template "STS#0" -TimeZone 4 -Wait #-RemoveDeletedSite
    }
}
while ($job.JobStateInfo -eq 'Running'){
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 5 
}
Write-Host "." -NoNewline

Write-Host "Importing Site Artifacts"
.\Provision-InfrastructureSiteArtifacts.ps1 -InfrastructureSiteUrl $InfrastructureSiteUrl  -AzureWebSiteUrl "http://$($azureService).azurewebsites.net" -Credentials $Credentials
Write-Host "Infrastructure Site Created. Url $InfrastructureSiteUrl"