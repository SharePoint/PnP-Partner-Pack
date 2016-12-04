[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Azure Data Center to host resource.")]
    $Location,
    [Parameter(Mandatory = $false, HelpMessage="Resource Group  Name. ")]
    $Name
)
write-host "Create-ResourceGroup.ps1 -Name $Name -Location $Location" -ForegroundColor Yellow

$name = ./Confirm-ParameterValue.ps1 "Confirm your Resource Group Name" -value $name
if($null -eq (Get-AzureRmResourceGroup -Name $name -Location $Location -ErrorAction 0 ) )
{
	New-AzureRmResourceGroup -Name $name -Location $Location | Out-Null
}
return $name