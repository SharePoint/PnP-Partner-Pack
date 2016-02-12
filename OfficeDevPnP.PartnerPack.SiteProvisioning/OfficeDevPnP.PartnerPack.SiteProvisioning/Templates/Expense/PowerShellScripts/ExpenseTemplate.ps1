<#
.SYNOPSIS
Uploads the template in the PnP PartnerPack Infra Site

.EXAMPLE
PS C:\> .\ExpenseTemplate.ps1 `
-PnpInfrasite "https://mytenant.sharepoint.com/sites/infrastructure" `
-ExpenseTemplate  "c:\Templates\ExpenseTemplate" 
#>

param
(
     [Parameter(Mandatory = $true, HelpMessage="Enter the Url of PnP Infrastructural site")]
    [String]
    $PnpInfrasite,

     [Parameter(Mandatory = $true, HelpMessage="Enter the path of Expense Template")]
    [String]
    $ExpenseTemplate

 )

Connect-SPOnline $PnpInfrasite
Add-SPOFile -Path $ExpenseTemplate -Folder "PnPProvisioningTemplates"