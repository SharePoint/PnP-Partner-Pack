<#
.SYNOPSIS
Uploads the template in the PnP PartnerPack Infra Site

.EXAMPLE
PS C:\> .\Expense.ps1 `
-PnpInfrasite "https://mytenant.sharepoint.com/sites/infrastructure" `
-Csspath  "c:\Templates\Expense.css" `
-Jspath "c:\Templates\Expense.min.js" `
-PeoplePickerCsspath  "c:\Templates\peoplepickercontrol.css" `
-PeoplePickerJspath "c:\Templates\peoplepickercontrol.js" `
-PeoplePickerAppJspath  "c:\Templates\app.js" `
-PeoplePickerResourceJspath "c:\Templates\peoplepickercontrol_resources.en.js" `
#>


param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the Url of PnP Infrastructural site")]
    [String]
    $PnpInfrasite,

    [Parameter(Mandatory = $true, HelpMessage="Enter the path of Expense.css")]
    [String]
    $Csspath,

    [Parameter(Mandatory = $true, HelpMessage="Enter the path of Expense.min.js")]
    [String]
    $Jspath,

	[Parameter(Mandatory = $true, HelpMessage="Enter the path of peoplepickercontrol.css")]
    [String]
    $PeoplePickerCsspath,

    [Parameter(Mandatory = $true, HelpMessage="Enter the path of peoplepickercontrol.js")]
    [String]
    $PeoplePickerJspath,

	[Parameter(Mandatory = $true, HelpMessage="Enter the path of app.js(peoplepickercontrol)")]
    [String]
    $PeoplePickerAppJspath,

	[Parameter(Mandatory = $true, HelpMessage="Enter the path of peoplepickercontrol_resources.en.js")]
    [String]
    $PeoplePickerResourceJspath

)

Connect-SPOnline $PnpInfrasite
Add-SPOFile -Path $Csspath -Folder "PnPProvisioningTemplates"
Add-SPOFile -Path $Jspath -Folder "PnPProvisioningTemplates"
Add-SPOFile -Path $PeoplePickerCsspath -Folder "PnPProvisioningTemplates"
Add-SPOFile -Path $PeoplePickerJspath -Folder "PnPProvisioningTemplates"
Add-SPOFile -Path $PeoplePickerAppJspath -Folder "PnPProvisioningTemplates"
Add-SPOFile -Path $PeoplePickerResourceJspath -Folder "PnPProvisioningTemplates"
