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
    [Parameter(Mandatory = $true, HelpMessage="App Service Plan name.")]
    $ServicePlan,
    [Parameter(Mandatory = $true, HelpMessage="Resource Group that holds all associated resources")]
    $ResourceGroupName,
    [Parameter(Mandatory = $true, HelpMessage="Certificate used for ssl binding.")]
    $CertificateFile,
    [Parameter(Mandatory = $true, HelpMessage="Certificate Common Name.")]
    $CertificateCommonName,
    [Parameter(Mandatory = $true, HelpMessage="Password of the certificate used for ssl binding.")]
     $CertificatePassword,
    [Parameter(Mandatory = $true, HelpMessage="Thumbprint of certificate used for ssl binding.")]
    $CertificateThumbprint
)
function CreateApp
{
    $app = (Get-AzureRmWebApp -Name $Name -ResourceGroupName $ResourceGroupName  -EA 0 )
    if($null -eq $app -or $app.Count -eq 0){
        try{
            $app=    New-AzureRmWebApp -Name $Name  -AppServicePlan $ServicePlan -ResourceGroupName $ResourceGroupName -Location $Location  
        }catch{
            write-error "Problems creating App Service Instance." -Exception $_.Exception 
            if("Y" -eq (read-host "Do You want to try again? [Y] for yes, anything else for no")){
               return  ./Create-AppService.ps1 -Name $name -ServicePlan $ServicePlan -ResourceGroupName $ResourceGroupName -CertificateFile $CertificateFile -CertificateCommonName $CertificateCommonName -CertificateThumbprint $CertificateThumbprint -CertificatePassword $CertificatePassword -Location            }else{
                break;
            }
        }
    }elseif($app.Count -gt 1){
        write-error "Multiple Applications found with the same name. Stopping now" 
        break;
    }
    return $app 

}
function SetProperties
{
    $hash = @{}
    ForEach ($s in $app.SiteConfig.AppSettings) {
        $hash[$s.Name] = $s.Value
    }
    $hash["WEBSITE_LOAD_CERTIFICATES"] = "*"
    $hash["WEBJOBS_IDLE_TIMEOUT"] = "10000"
    $hash["SCM_COMMAND_IDLE_TIMEOUT"] = "10000"

    Set-AzureRMWebApp -ResourceGroupName $ResourceGroupName -Name $Name -AppSettings $hash 

}
function SetSslCertificate{

if((Get-AzureRmWebAppSSLBinding -WebAppName $name -Name $CertificateCommonName -ResourceGroupName $ResourceGroupName).Count -eq 0){
    $pfxPath = resolve-path "./$($CertificateFile).pfx"
    $cerPath = resolve-path "./$($CertificateFile).cer"

try{
    ### from https://github.com/Azure/azure-powershell/issues/2108
    ## this is a bit of a hacky way of achieving the outcome, but it does work.
    New-AzureRmWebAppSSLBinding -ResourceGroupName $ResourceGroupName  `
                                -WebAppName $Name -CertificateFilePath  $pfxPath  `
                                -CertificatePassword $CertificatePassword  `
                                -Name $CertificateCommonName -ErrorAction Stop  -WarningAction SilentlyContinue   #Suppress the Warning on CNAME record
} catch {
        <# need to suppress the error of the SSL Binding - replace the cmdlet when changes#>
        $msg = $_
        $hostnamemsg = "Hostname '" + $CertificateCommonName + "' does not exist."
        if($msg.tostring() -eq $hostnamemsg.tostring()) {
             $ReturnValue = $True
        }else {
            write-host "Encountered error while uploading the certificate to the WebApp. Error Message is $msg." -ForegroundColor Red
        }
    } 
} 


}
write-host "Create-AppService.ps1  -Name $Name  -Location $Location -ResourceGroupName $ResourceGroupName " -ForegroundColor Yellow

$name = ./Confirm-ParameterValue.ps1 "Confirm your App Service/Web App name" -value $name

$app = CreateApp
SetProperties 

SetSSLCertificate 
return $name