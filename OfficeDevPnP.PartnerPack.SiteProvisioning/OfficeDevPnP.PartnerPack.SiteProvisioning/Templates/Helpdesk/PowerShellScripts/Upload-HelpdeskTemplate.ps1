<#
.SYNOPSIS
Uploads the template in the PnP PartnerPack Infra Site

.EXAMPLE
PS C:\> .\HelpdeskTemplate.ps1 `
-PnpInfrasite "https://mytenant.sharepoint.com/sites/infrastructure" `
-HelpdeskTemplate  "c:\Templates\HelpdeskTemplate" 
#>

param
(
     [Parameter(Mandatory = $true, HelpMessage="Enter the Url of PnP Infrastructural site")]
    [String]
    $PnpInfrasite,

     [Parameter(Mandatory = $true, HelpMessage="Enter the path of Helpdesk Template")]
    [String]
    $HelpdeskTemplate

 )

 Connect-SPOnline $PnpInfrasite
Add-SPOFile -Path $HelpdeskTemplate -Folder "PnPProvisioningTemplates"