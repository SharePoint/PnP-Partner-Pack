param($force = $true)
if($Global:config -eq $null -or $force )
{
    $Global:config = @{
        # Update with the name of your subscription.
        SubscriptionName = "Contoso Production"

        #Office 365 Tenant name(ie: 'contoso' if tenant is contoso.sharepoint.com)
        Tenant = "contoso"

        #Full URL of where the partner pack is to be deployed
        InfrastructureSiteUrl ="https://contoso.sharepoint.com/sites/PnP-Partner-Pack-Infrastructure"
        
        #Primary Site Collection Owner
        InfrastructureSiteOwner =$null 

        # Give a name to your new storage account. It MUST be lowercase!
        #storage account is not created should there be a default account already in polace
        StorageAccountName = "pnppartnerpackstorage9"
        StorageAccountSku ="Standard_LRS"
        #Azure datacenter location. prefer using lower case
        Location = $null
        #resource group 
        ResourceGroupName= "PnPPartnerPack" 
        #app service tier. Partner Pack requires at least basic(due to SSL Certificates)
        AppServiceTier="Basic" 

        #app Service Name. This will also be used for the URL of the solution ie(contoso.azurewebsites.net)
        AppServiceName = "OfficeDevPnPPartnerPackSiteProvisioning9"

        #App service plan name. 
        AppServicePlanName = "PnPPartnerPack"

        #Self Signed Certificate common name.(cn=$CertificateCommonName ). Name is irrelevant as certificate isn't bound to HTTPS requests.
        CertificateCommonName = "contoso.com"
        #password to be used in certificate. 
        CertificatePassword = ./New-AppSecret.ps1 -length 10 
        # randomly generated app secret. 
        AppClientSecret=./New-AppSecret.ps1 
        #ApplicationIdentifierUri is identifier used by Azure AD Application
        ApplicationIdentifierUri = $null 
    }
}

$Global:config  