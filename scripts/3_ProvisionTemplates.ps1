# Provision Contents
#
#
# Provide URL to the infrastructure site. Site needs to be created. For example: https://yourtenant.sharepoint.com/sites/infrastructure
$InfrastructureSiteUrl = ""
# Provide URL to the azureweb site hosting the app. For example: https://yourprovisioningapp.azurewebsites.net
$AzureWebSiteUrl = ""

# DO NOT MODIFY BELOW

while($InfrastructureSiteUrl -eq "" -or $InfrastructureSiteUrl -eq $null -or $InfrastructureSiteUrl.ToLower() -notlike "https://*")
{
    $InfrastructureSiteUrl = Read-Host -Prompt "Enter infrastructure site url (e.g. https://yourtenant.sharepoint.com/sites/infrastructure)"
}

while($AzureWebSiteUrl -eq "" -or $AzureWebSiteUrl -eq $null -or $AzureWebSiteUrl.ToLower() -notlike "https://*")
{
    $AzureWebSiteUrl = Read-Host -Prompt "Enter azure web site url (e.g. https://yourprovisioningapp.azurewebsites.net)"
}

Write-Host -ForegroundColor Yellow "Modifying Templates"
$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning\OfficeDevPnP.PartnerPack.SiteProvisioning\templates"

# CUSTOMBAR
$provisioningTemplate = (Get-Content "$basePath\CustomBar\SPO-CustomBar-Base.xml") -As [Xml]

$customActions = $provisioningTemplate.Provisioning.Templates.ProvisioningTemplate.CustomActions.SiteCustomActions.CustomAction
$spoCustomBar = $customActions | ?{$_.Name -eq "spoCustomBar"}
$scriptBlock = $spoCustomBar.ScriptBlock.Replace("[AZUREWEBSITE]",$AzureWebSiteUrl)
$spoCustomBar.ScriptBlock = $scriptBlock

$composedLook = $provisioningTemplate.Provisioning.Templates.ProvisioningTemplate.ComposedLook
$composedLook.AlternateCSS = $composedLook.AlternateCSS.Replace("[AZUREWEBSITE]",$AzureWebSiteUrl)

$provisioningTemplate.Save("$basePath\CustomBar\SPO-CustomBar.xml")

# OVERRIDES
$provisioningTemplate = (Get-Content "$basePath\Overrides\PnP-Partner-Pack-Overrides-Base.xml") -As [Xml]

$customActions = $provisioningTemplate.Provisioning.Templates.ProvisioningTemplate.CustomActions.SiteCustomActions.CustomAction
$overridesAction = $customActions | ?{$_.Name -eq "PnPPartnerPackOverrides"}
$scriptBlock = $overridesAction.ScriptBlock.Replace("[AZUREWEBSITE]",$AzureWebSiteUrl)
$overridesAction.ScriptBlock = $scriptBlock

$provisioningTemplate.Save("$basePath\Overrides\PnP-Partner-Pack-Overrides.xml")

# RESPONSIVE
$provisioningTemplate = (Get-Content "$basePath\Responsive\SPO-Responsive-Base.xml") -As [Xml]

$customActions = $provisioningTemplate.Provisioning.Templates.ProvisioningTemplate.CustomActions.SiteCustomActions.CustomAction
$spoResponsive = $customActions | ?{$_.Name -eq "spoResponsive"}
$scriptBlock = $spoResponsive.ScriptBlock.Replace("[AZUREWEBSITE]",$AzureWebSiteUrl)
$spoResponsive.ScriptBlock = $scriptBlock

$composedLook = $provisioningTemplate.Provisioning.Templates.ProvisioningTemplate.ComposedLook
$composedLook.AlternateCSS = $composedLook.AlternateCSS.Replace("[AZUREWEBSITE]",$AzureWebSiteUrl)

$provisioningTemplate.Save("$basePath\Responsive\SPO-Responsive.xml")

Write-Host -ForegroundColor Yellow "Provisioning Templates"
Connect-SPOnline $InfrastructureSiteUrl
Apply-SPOProvisioningTemplate $basePath\PnP-Partner-Pack-Infrastructure-Contents.xml