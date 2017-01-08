<#
.SYNOPSIS
Creates the Azure App Service that will host the PnP Partner Pack web application and web jobs

.DESCRIPTION


.EXAMPLE
PS C:\> ./Create-AppService.ps1  -Location "australia southeast" `
                            -Name  "PnPPartnerPackAppService" `
                            -ServicePlan "PnPPartnerPackAppPlan" `
                            -ResourceGroupName "PnPPartnerPack" `
                            -AppServiceTier "Basic" `
                            -CertificateCommonName "contoso.com" `
                            -CertificateFile "contoso.com" `
                            -CertificatePassword "Password1" 
                            -CertificateThumbprint "61402FFFEB61CA27FABB946193D88A7136ED2310"

#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Azure Data Center to host resource.")]
    $Location,
    [Parameter(Mandatory = $false, HelpMessage="App Service Name. ")]
    $Name,
    [Parameter(Mandatory = $true, HelpMessage="Resource Group that holds all associated resources")]
    $ResourceGroupName,
    [Parameter(Mandatory = $false, HelpMessage="App Service Tier. You need at least basic for this solution")]
    $AppServiceTier="Basic"     
)
write-host "Create-AppServicePlan.ps1 -Name $Name -Location $Location -ResourceGroupName $ResourceGroupName -AppServiceTier $AppServiceTier" -ForegroundColor Yellow

$name = ./Confirm-ParameterValue.ps1 -prompt "Confirm the App Service Plan name" -value $name
$AppServiceTier = ./Confirm-ParameterValue.ps1 -prompt "Confirm the Tier " -value $AppServiceTier -options ("Basic", "Standard","Premium")

if($null -eq (Get-AzureRmAppServicePlan -Name $Name -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue))
{
    New-AzureRmAppServicePlan -Location $Location -ResourceGroupName $ResourceGroupName -Name $name -Tier $AppServiceTier | Out-Null
}else { 
	if([String]::IsNullOrEmpty( (Read-Host "Service Plan already exists. Do you want to use it? [Enter] for Yes, type anything for No"))){
        Set-AzureRmAppServicePlan -Name $Name -Tier $AppServiceTier -ResourceGroupName $ResourceGroupName |Out-Null
    }else {
        return ./Create-AppServicePlan.ps1 -Location $location -Name $Name -ResourceGroupName $ResourceGroupName -AppServiceTier $AppServiceTier
    }
}
return @{
    Name = $Name
    Tier = $AppServiceTier
}
