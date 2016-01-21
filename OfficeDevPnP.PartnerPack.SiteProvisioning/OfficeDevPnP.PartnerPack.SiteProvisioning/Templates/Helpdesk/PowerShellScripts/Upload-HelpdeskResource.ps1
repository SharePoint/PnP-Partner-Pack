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
