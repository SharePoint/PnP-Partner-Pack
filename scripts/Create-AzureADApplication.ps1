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
    $KeyCredentials
)

write-host "Create-AzureADApplication.ps1 -ApplicationServiceName $ApplicationServiceName -ApplicationServiceName $ApplicationServiceName -ApplicationIdentifierUri $ApplicationIdentifierUri -AppServiceTier $AppServiceTier" -ForegroundColor Yellow

function SetKeys
{
    ##### 2016-11-07 20-08 Azure AD API is behaving oddly. Some quests will throw an exception when adding an artefact shortly after removing one of its kind.
    #### both New-AzureADApplicationKeyCredential and New-AzureADApplicationPasswordCredential throw the following exception
    #### New-AzureADApplicationKeyCredential : Error occurred while executing SetApplication 
    #### StatusCode: BadRequest
    #### ErrorCode: Request_BadRequest
    #### Message: A stream property was found in a JSON Light request payload. Stream properties are only supported in responses.

    #### it seems that issue is mitigated by adding Sleep calls 

    $enc = [System.Text.Encoding]::ASCII
    $KeyIdentifier = "pnppartnerpack"
    Get-AzureADApplicationKeyCredential -ObjectId $app.ObjectId | 
                                    Where-Object { 
                                        $null -ne $_.CustomKeyIdentifier -and $enc.GetString($_.CustomKeyIdentifier) -eq $KeyIdentifier 
                                    }| ForEach-Object {
                                       Remove-AzureADApplicationKeyCredential -ObjectId $app.ObjectId -KeyId $_.KeyId
                                       Sleep -Seconds 10
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
                                        Remove-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId -KeyId $_.KeyId
                                        Sleep -Seconds 10
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
#$certificateInfo =  ./GEt-SelfSignedCertificateInformation.ps1 -CertificateFile $CertificateFile 
$app = ((Get-AzureRmADApplication) | Where-Object { $_.IdentifierUris -contains $ApplicationIdentifierUri.ToString()} )
if($null -eq $app){
    $app = New-AzureRmADApplication -DisplayName $ApplicationServiceName  -HomePage $homepage -IdentifierUris $ApplicationIdentifierUri.ToLower() -CertValue $KeyCredentials.value # $.$key 
}

$clientSecret = SetKeys
$replyUrls = @($homepage.ToLower())
$replyUrls += "https://localhost:44300/" 
Set-AzureRmADApplication -ObjectId $app.ObjectId  -ReplyUrls $replyUrls

SetServicePrincipal
return @{
ClientSecret = $ClientSecret 
AADApp = $app

}