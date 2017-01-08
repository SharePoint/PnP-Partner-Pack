<#
.SYNOPSIS
Creates the Azure App Service that will host the PnP Partner Pack web application and web jobs

.DESCRIPTION
Creates the Azure App Service that will host the PnP Partner Pack web application and web jobs.
This also adds a Key Credential and a password credential(AppSecret) as well as register the service principal and grant permissions


.EXAMPLE
PS C:\> ./Create-AzureADApplication.ps1  -Tenant "contoso" `
                            -ApplicationServiceName  "PnPPartnerPackAppService" `
                            -ApplicationIdentifierUri "https://contoso.onmicrosoft.com/PnPPartnerPackAppService.azurewebsites.net" `
                            -CertificateFile "contoso.com" 
Adds a new Azure AD Application. Sets 
            display name to "PnPPartnerPackAppService", 
            homepage to "https://PnPPartnerPackAppService.azurewebsites.net/"


#>


[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Office 365 Tenant name.")]
    $Tenant,
    [Parameter(Mandatory = $true, HelpMessage="App ServiceName.")]
    $ApplicationServiceName,

    [Parameter(Mandatory = $true, HelpMessage="App Service Identifier Uri. ")]
    $ApplicationIdentifierUri,
  
    [Parameter(Mandatory = $true, HelpMessage="Key Credentials to be used by Azure Web App.")]
    $KeyCredentials, 

    [Parameter(Mandatory = $false, HelpMessage="Add http://localhost:44300 as reply Url if it is true.")]
    $LocalDebug = $false
)
write-host "Create-AzureADApplication.ps1 -ApplicationServiceName $ApplicationServiceName -ApplicationServiceName $ApplicationServiceName -ApplicationIdentifierUri $ApplicationIdentifierUri -AppServiceTier $AppServiceTier" -ForegroundColor Yellow

function SetKeys
{
    $enc = [System.Text.Encoding]::ASCII
    $KeyIdentifier = "pnppartnerpack"
    Get-AzureADApplicationKeyCredential -ObjectId $app.ObjectId | 
                                    Where-Object { 
                                        $null -ne $_.CustomKeyIdentifier -and $enc.GetString($_.CustomKeyIdentifier) -eq $KeyCredentials.customKeyIdentifier 
                                    }| ForEach-Object {
                                        $key = $_
                                        Remove-AzureADApplicationKeyCredential -ObjectId $app.ObjectId -KeyId $key.KeyId
                                        Sleep -Seconds 2
                                    }
    New-AzureADApplicationKeyCredential -ObjectId $app.ObjectId `
                                    -CustomKeyIdentifier $KeyCredentials.customKeyIdentifier `
                                    -StartDate (Get-DAte).ToUniversalTime() `
                                    -EndDate (get-Date).AddYears(2) -Usage Verify `
                                    -Value $KeyCredentials.value -Type AsymmetricX509Cert |Out-Null
    Sleep -Seconds 2
    Get-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId | 
                                    Where-Object { 
                                        $enc.GetString($_.CustomKeyIdentifier) -eq $keyIdentifier 
                                    } | ForEach-Object {
                                        $passwordCredential = $_
                                        Remove-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId -KeyId $passwordCredential.KeyId
                                        Sleep -Seconds 2
                                    }
    $passwordCredential = New-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId -CustomKeyIdentifier $keyIdentifier 
    return $passwordCredential.Value                                    
}

function SetServicePrincipal
{
    if($null -eq (GEt-AzureRmADServicePrincipal -SearchString $app.DisplayName)){
        Sleep -Seconds 20 
        New-AzureRmADServicePrincipal -ApplicationId $app.ApplicationId |Out-Null 
        Sleep -Seconds 5
        New-AzureRmRoleAssignment -RoleDefinitionName Reader -ServicePrincipalName $app.ApplicationId |Out-Null
    }
}

$homepage = "https://$ApplicationServiceName.azurewebsites.net/".ToLower()
$app = ((Get-AzureRmADApplication) | Where-Object { $_.IdentifierUris -contains $ApplicationIdentifierUri.ToString()} )
if($null -eq $app){
    $app = New-AzureRmADApplication -DisplayName $ApplicationServiceName  -HomePage $homepage -IdentifierUris $ApplicationIdentifierUri.ToLower() -CertValue $KeyCredentials.value # $.$key 
}

$clientSecret = SetKeys
$replyUrls = @($homepage.ToLower())
if($LocalDebug){
    $replyUrls += "https://localhost:44300/" 
}
Set-AzureRmADApplication -ObjectId $app.ObjectId  -ReplyUrls $replyUrls
SetServicePrincipal

return @{
    ClientSecret = $clientSecret 
    AADApp = $app

}