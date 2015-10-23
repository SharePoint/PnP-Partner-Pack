Param(

   [Parameter(Mandatory=$true)]
   [string]$Name,

   [Parameter(Mandatory=$true)]
   [ValidateSet('South Central US','North Europe','Web Europe','East US','East Asia','Southeast Asia','West US','Central US','Japan West','Japan East','North Central US','East US 2','Brazil South')]
   [string]$Location,

   [Parameter(Mandatory=$true)]
   [string]$ClientId,

   [Parameter(Mandatory=$true)]
   [string]$AzureTenant,

   [Parameter(Mandatory=$true)]
   [string]$CertificatePath,

   [Parameter(Mandatory=$true)]
   [string]$CertificatePassword
)


# Do not modify below
function CreateCertificate
{
    $certroot = [System.IO.Path]::GetTempPath();
    $certificationname = "PnPPartnerPackCertificate";
    $password = read-host "Enter certificate password" -AsSecureString
    $startdate = Get-Date # now
    $enddate = $startdate.AddYears(2) # validate the certificate for 2 years
    New-SelfSignedCertificate -KeyLength 2048 -HardwareKeyUsage EncryptionKey -KeyUsage 
}


Write-Progress -Activity "Creating Site" -PercentComplete 0
Write-Progress -Activity "Creating Site" -Status "Creating new azure site" -PercentComplete 10

$site = New-AzureWebsite -Location $Location -Name $Name

$connectionStrings = ( @{Name = "CertificatePassword"; Type = "Custom"; ConnectionString = $CertificatePassword} )

$settings = @{"CertPath" = $CertificatePath;  
              "AzureTenant" = $AzureTenant;
              "ClientId" = $ClientId;
             }

Write-Progress -Activity "Creating Site" -Status "Setting appsettings and connection strings" -PercentComplete 20
$site | Set-AzureWebSite -ConnectionStrings $connectionStrings -AppSettings $settings

# Uploading Job

$jobCollectionName = "PartnerPack-JobCollection"
$jobName = "EnforceAdministratorsJob"

Write-Progress -Activity "Uploading Enforce Administrators Job" -PercentComplete 30
$job = $site | New-AzureWebsiteJob -JobName $jobName -JobType Triggered -JobFile ..\OfficeDevPnP.PartnerPack.WebJobs\AdministrationJob.zip
$jobCollection = Get-AzureSchedulerJobCollection -Location $Location -JobCollectionName $jobCollectionName -ErrorAction SilentlyContinue
if($jobCollection -eq $null)
{
    $jobCollection = New-AzureSchedulerJobCollection -Location $Location -JobCollectionName $jobCollectionName;
}

$authPair = "$($site.PublishingUsername):$($site.PublishingPassword)";
$pairBytes = [System.Text.Encoding]::UTF8.GetBytes($authPair);
$encodedPair = [System.Convert]::ToBase64String($pairBytes);
Write-Progress -Activity "Uploading Enforce Administrators Job" -Status "Creating daily schedule" -PercentComplete 40
$schedulerJob = Get-AzureSchedulerJob -Location $Location -JobCollectionName $jobCollectionName -JobName "EnforceAdministratorsJob" -ErrorAction SilentlyContinue
if($schedulerJob -ne $null)
{
    Remove-AzureSchedulerJob -Location $Location -JobCollectionName $jobCollectionName -JobName $jobName -Force
}
New-AzureSchedulerHttpJob -JobCollectionName $jobCollectionName -JobName $jobName -Method POST -URI "$($job.Url)\run" -Location $Location -StartTime "2015-01-01" -Interval 1 -Frequency Day -Headers @{ "Content-Type" = "text/plain"; "Authorization" = "Basic $encodedPair"; };
