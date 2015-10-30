# Provision Contents
#
#
# Provide URL to the infrastructure site. Site needs to be created. For example: https://yourtenant.sharepoint.com/sites/infrastructure
$InfrastructureSiteUrl = ""

# DO NOT MODIFY BELOW
while($InfrastructureSiteUrl -eq "" -or $InfrastructureSiteUrl -eq $null)
{
    $InfrastructureSiteUrl = Read-Host -Prompt "Enter infrastructure site URL"
}

$creds = Get-Credential -Message "Enter credentials for the infrastructure site"
Connect-SPOnline $InfrastructureSiteUrl -Credentials $creds
Apply-SPOProvisioningTemplate $basePath\PnP-Partner-Pack-Infrastructure-Contents.xml