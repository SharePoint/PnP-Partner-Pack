<#
.SYNOPSIS
Uploads the template in the PnP PartnerPack Infra Site

.EXAMPLE
PS C:\> .\Contact.ps1 `
-PnpInfrasite "https://mytenant.sharepoint.com/sites/infrastructure" `
-Csspath  "c:\Templates\Contact.css" `
-Jspath "c:\Templates\Contact.min.js" `
#>




param
(
     [Parameter(Mandatory = $true, HelpMessage="Enter the Url of PnP Infrastructural site")]
    [String]
    $PnpInfrasite,

    [Parameter(Mandatory = $true, HelpMessage="Enter the path of contact.css")]
    [String]
    $Csspath,

    [Parameter(Mandatory = $true, HelpMessage="Enter the path of contact.min.js")]
    [String]
    $Jspath

)

Connect-SPOnline $PnpInfrasite
Add-SPOFile -Path $Csspath -Folder "PnPProvisioningTemplates"
Add-SPOFile -Path $Jspath -Folder "PnPProvisioningTemplates"
