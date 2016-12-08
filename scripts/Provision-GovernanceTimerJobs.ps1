<#
.SYNOPSIS
Uploads and configures the governance timerjobs.

.DESCRIPTION
It assumes that the app.config and web.config files have been configured correctly.

.EXAMPLE
PS C:\> .\Provision-GovernanceTimerJobs.ps1 -Location "North Europe" -AzureWebSite WebsiteName

This will pack the 2 governance webjobs, it will upload them to the azurewebsite specified, and it will create a default schedule to run the 
2 jobs every day at 05:00 AM GMT.

.EXAMPLE
PS C:\> .\Provision-GovernanceTimerJobs.ps1 -Location "North Europe" -AzureWebSite WebsiteName - StartDateAndTime "2015-11-08 09:00"

.EXAMPLE
Use the Get-AzureSubscription in combination with Get-AzureWebsite to locate which Azure Subscription hosts your website

PS C:\> .\Provision-GovernanceTimerJobs.ps1 -Location "North Europe" -AzureWebSite WebsiteName -AzureSubscriptionName "Visual Studio Enterprise"

#>
Param(

   [Parameter(Mandatory=$false)]
   [ValidateSet('Debug','Release')]
   [string]$Build = "Release",

   [Parameter(Mandatory=$true)]
   [ValidateSet('South Central US','North Europe','Web Europe','East US','East Asia','Southeast Asia','West US','Central US','Japan West','Japan East','North Central US','East US 2','Brazil South')]
   [string]$Location,

   [Parameter(Mandatory=$true)]
   [string]$AzureWebSite,
   
   [Parameter(Mandatory=$true)]
   [ValidateSet('Access to Azure Active Directory', 'Developer Program Benefit', 'Visual Studio Enterprise')]
   [string]$AzureSubscriptionName = "Access to Azure Active Directory",

   [Parameter(Mandatory=$false)]
   [string]$JobCollectionName = "PnP-PartnerPack-GovernanceJobCollection",

   [Parameter(Mandatory=$false)]
   [string]$StartDateAndTime = (Get-Date -Format "yyyy-MM-dd 06:00")
)

# DO NOT MODIFY BELOW

Add-AzureAccount

Add-Type -AssemblyName System.IO.Compression.FileSystem

Select-AzureSubscription -SubscriptionName $AzureSubscriptionName

function UploadJob($jobName,$jobFile)
{
    Write-Host -ForegroundColor Yellow "Uploading $jobName"
    $site = Get-AzureWebsite -Name $AzureWebSite
    $job = $site | New-AzureWebsiteJob -JobName $jobName -JobType Triggered -JobFile $jobFile
    $jobCollection = Get-AzureSchedulerJobCollection -Location $Location -JobCollectionName $JobCollectionName -ErrorAction SilentlyContinue
    if($jobCollection -eq $null)
    {
        $jobCollection = New-AzureSchedulerJobCollection -Location $Location -JobCollectionName $JobCollectionName;
    }

    $authPair = "$($site.PublishingUsername):$($site.PublishingPassword)";
    $pairBytes = [System.Text.Encoding]::UTF8.GetBytes($authPair);
    $encodedPair = [System.Convert]::ToBase64String($pairBytes);
    $schedulerJob = Get-AzureSchedulerJob -Location $Location -JobCollectionName $jobCollectionName -JobName $jobName -ErrorAction SilentlyContinue
    if($schedulerJob -ne $null)
    {
        Remove-AzureSchedulerJob -Location $Location -JobCollectionName $jobCollectionName -JobName $jobName -Force
    }
    #New-AzureSchedulerHttpJob -JobCollectionName $jobCollectionName -JobName $jobName -Method POST -URI "$($job.Url)\run" -Location $Location -StartTime $StartDateAndTime -Interval 1 -Frequency Day -Headers @{ "Content-Type" = "text/plain"; "Authorization" = "Basic $encodedPair"; };
    New-AzureSchedulerHttpJob -JobCollectionName $jobCollectionName -JobName $jobName -Method POST -URI "$($job.Url)\run" -Location $Location -StartTime $StartDateAndTime -Interval 1 -Frequency Day
}

$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning"

Write-Host -ForegroundColor Yellow "Packing Enforce Administrators Job"

# Check if the required release files are present
$files = Get-ChildItem "$basePath\OfficeDevPnP.PartnerPack.CheckAdminsJob\bin\$Build" -ErrorAction SilentlyContinue
if($files -ne $null -and $files.Length -gt 0)
{
    # Pack the files into a ZIP file
    $zipFile = Get-ChildItem "$basePath\OfficeDevPnP.PartnerPack.CheckAdminsJob\EnforceAdminsJob.zip" -ErrorAction SilentlyContinue
    if($zipFile -ne $null)
    {
        Remove-Item "$basePath\OfficeDevPnP.PartnerPack.CheckAdminsJob\EnforceAdminsJob.zip"
    }
    [IO.Compression.ZipFile]::CreateFromDirectory("$basePath\OfficeDevPnP.PartnerPack.CheckAdminsJob\bin\$Build","$basePath\OfficeDevPnP.PartnerPack.CheckAdminsJob\EnforceAdminsJob.zip");
    UploadJob "EnforceTwoAdminsJob" "$basePath\OfficeDevPnP.PartnerPack.CheckAdminsJob\EnforceAdminsJob.zip"
} else {
    Write-Host -ForegroundColor Cyan "No build files available. Please configure and build the solution first."
}

Write-Host -ForegroundColor Yellow "Packing Check External Users Job"
$files = Get-ChildItem "$basePath\OfficeDevPnP.PartnerPack.ExternalUsersJob\bin\$Build" -ErrorAction SilentlyContinue
if($files -ne $null -and $files.Length -gt 0)
{
    # Pack the files into a ZIP file
    $zipFile = Get-ChildItem "$basePath\OfficeDevPnP.PartnerPack.ExternalUsersJob\ExternalUsersJob.zip" -ErrorAction SilentlyContinue
    if($zipFile -ne $null)
    {
        Remove-Item "$basePath\OfficeDevPnP.PartnerPack.CheckAdminsJob\ExternalUsersJob.zip"
    }
    [IO.Compression.ZipFile]::CreateFromDirectory("$basePath\OfficeDevPnP.PartnerPack.ExternalUsersJob\bin\$Build","$basePath\OfficeDevPnP.PartnerPack.ExternalUsersJob\ExternalUsersJob.zip");
    UploadJob "ExternalUsersJob" "$basePath\OfficeDevPnP.PartnerPack.ExternalUsersJob\ExternalUsersJob.zip"
} else {
    Write-Host -ForegroundColor Cyan "No build files available. Please configure and build the solution first."
}
