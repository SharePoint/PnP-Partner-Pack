<#
.SYNOPSIS
Uploads the template in the PnP PartnerPack Infra Site

.EXAMPLE
PS C:\> .\ContactTemplate.ps1 `
-PnpInfrasite "https://mytenant.sharepoint.com/sites/infrastructure" `
-ContactTemplate  "c:\Templates\ContactTemplate" 
#>


param
(
     [Parameter(Mandatory = $true, HelpMessage="Enter the Url of PnP Infrastructural site")]
    [String]
    $PnpInfrasite,

     [Parameter(Mandatory = $true, HelpMessage="Enter the path of Contact Template")]
    [String]
    $ContactTemplate

 )

 Connect-SPOnline $PnpInfrasite
Add-SPOFile -Path $ContactTemplate -Folder "PnPProvisioningTemplates"