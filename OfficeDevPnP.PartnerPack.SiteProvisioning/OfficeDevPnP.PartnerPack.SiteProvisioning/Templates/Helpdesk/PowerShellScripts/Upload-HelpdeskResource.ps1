<#
.SYNOPSIS
Uploads the template in the PnP PartnerPack Infra Site

.EXAMPLE
PS C:\> .\Helpdesk.ps1 `
-PnpInfrasite "https://mytenant.sharepoint.com/sites/infrastructure" `
-Csspath  "c:\Templates\Helpdesk.css" `
-Jspath "c:\Templates\Helpdesk.min.js" `
#>


param
(
     [Parameter(Mandatory = $true, HelpMessage="Enter the Url of PnP Infrastructural site")]
    [String]
    $PnpInfrasite,

    [Parameter(Mandatory = $true, HelpMessage="Enter the path of Helpdesk.css")]
    [String]
    $Csspath,

    [Parameter(Mandatory = $true, HelpMessage="Enter the path of Helpdesk.min.js")]
    [String]
    $Jspath

)

Connect-SPOnline $PnpInfrasite
Add-SPOFile -Path $Csspath -Folder "PnPProvisioningTemplates"
Add-SPOFile -Path $Jspath -Folder "PnPProvisioningTemplates"
