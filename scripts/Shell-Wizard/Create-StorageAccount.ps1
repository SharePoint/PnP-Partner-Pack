<#
.SYNOPSIS
Configures the Azure storage account 

.DESCRIPTION
This script creates an azure storage account should there be no storage accounts configured in your selected subscription

.EXAMPLE
PS C:\> .\Create-StorageAccount.ps1  -name "AzureStorageAccount" -ResourceGroupName  "PnPPartnerPack" -location "australia southeast" -subscriptionName "Contoso Production"

This will create a storage account "AzureStorageAccount" and associate it as the default storage account for subscription contoso production 

#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Azure Data Center to host resource.")]
    $Location,
    [Parameter(Mandatory = $false, HelpMessage="Storage Account Name. ")]
    $Name,
    [Parameter(Mandatory = $false, HelpMessage="Storage Account Sku. ")]
    $Sku,    
    [Parameter(Mandatory = $true, HelpMessage="Azure Subscription name.")]
    $SubscriptionName,
    [Parameter(Mandatory = $true, HelpMessage="Resource Group that holds all associated resources")]
    $ResourceGroupName
)
write-host "Create-StorageAccount.ps1 -Name $name -Sku $Sku -SubscriptionName $SubscriptionName" -ForegroundColor Yellow
$subscription = Get-AzureRmSubscription -SubscriptionName $SubscriptionName 
if($null -ne $subscription.CurrentStorageAccountName){
    write-host "Storage Account Already configured for this subscription. the current storage account is $($subscription.GetAccountName())" -ForegroundColor Red 
    if($subscription.GetAccountName() -eq $Name){
        $useExisting = Read-Host -Prompt "Do you want to use this storage account?(Y/N)"
        if((-not [String]::IsNullOrEmpty($useExisting)) -and $useExisting.ToUpper() -eq "Y"){
            return Get-AzureStorageKey -StorageAccountName $subscription.GetAccountName() 
        }
    }
}

$name = ./Confirm-ParameterValue.ps1 "Name your Storage Account" -value $Name
$Sku = ./Confirm-ParameterValue.ps1 "Name your Storage Account Sku" -value $sku
try{
    if($null -eq (get-AzureRmStorageAccount -name $name -ResourceGroupName $ResourceGroupName  -ErrorAction SilentlyContinue)){
        New-AzureRMStorageAccount -Name $Name.ToLower() -Location $Location -ResourceGroupName $ResourceGroupName -SkuName $Sku | Out-Null  
    }
}catch{
    write-error -Message "Exception processing storage account. The message is" -Exception $_.Exception
    $read =read-host -prompt "Do you want to try again with a different name?[Y] for yes, anything else for no" 
    if((-not [String]::IsNullOrEmpty($read)) -and $read.ToUpper() -eq "Y"){
        return ./Create-StorageAccount.ps1 -Location $location -Name $newAccountName -ResourceGroupName $ResourceGroupName
    }
    throw
}
return @{ 
    Name = $name 
    Sku = $Sku
    Key = ((Get-AzureRmStorageAccountKey -Name $name -ResourceGroupName $ResourceGroupName)| Where-Object{ $_.KeyName -eq "Key1"}).Value
}