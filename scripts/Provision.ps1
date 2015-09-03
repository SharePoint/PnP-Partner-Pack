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
Write-Progress -Activity "Uploading Enforce Administrators Job" -PercentComplete 30
$job = $site | New-AzureWebsiteJob -JobName "EnforceAdministratorsJob" -JobType Triggered -JobFile ..\OfficeDevPnP.PartnerPack.WebJobs\AdministrationJob.zip
$jobCollection = New-AzureSchedulerJobCollection -Location $Location -JobCollectionName "$($name)-job-collection";
$authPair = "$($site.PublishingUsername):$($site.PublishingPassword)";
$pairBytes = [System.Text.Encoding]::UTF8.GetBytes($authPair);
$encodedPair = [System.Convert]::ToBase64String($pairBytes);
Write-Progress -Activity "Uploading Enforce Administrators Job" -Status "Creating daily schedule" -PercentComplete 40
New-AzureSchedulerHttpJob -JobCollectionName $jobCollection[0].JobCollectionName -JobName "EnforceAdministratorsJob" -Method POST -URI "$($job.Url)\run" -Location $location -StartTime "2015-01-01" -Interval 1 -Frequency Day -Headers @{ "Content-Type" = "text/plain"; "Authorization" = "Basic $encodedPair"; };

