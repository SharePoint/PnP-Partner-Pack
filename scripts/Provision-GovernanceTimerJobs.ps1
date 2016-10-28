<#
.SYNOPSIS
Uploads and configures the governance timerjobs.

.DESCRIPTION
It assumes that the app.config and web.config files have been configured correctly.

.EXAMPLE
PS C:\> .\Provision-GovernanceTimerJobs.ps1 -Location "North Europe" -AzureWebSite yourwebsite

This will pack the 2 governance webjobs, it will upload them to the azurewebsite specified, and it will create a default schedule to run the 
2 jobs every day at 05:00 AM GMT.

.EXAMPLE
PS C:\> .\Provision-GovernanceTimerJobs.ps1 -Location "North Europe" -AzureWebSite yourwebsite - StartDateAndTime "2015-11-08 09:00"

#>
Param(

   [Parameter(Mandatory=$false)]
   [ValidateSet('Debug','Release')]
   [string]$Build = "Release",

   [Parameter(Mandatory=$true)]
   [ValidateSet('South Central US','North Europe','Web Europe','East US','East Asia','Southeast Asia','West US','Central US','Japan West','Japan East','North Central US','East US 2','Brazil South','Australia East', 'Australia Southeast')]
   [string]$Location,

   [Parameter(Mandatory=$true)]
   [string]$AzureWebSite,

   [Parameter(Mandatory=$false)]
   [string]$JobCollectionName = "PnP-PartnerPack-GovernanceJobCollection",

   [Parameter(Mandatory=$false)]
   [string]$StartDateAndTime = (Get-Date -Format "yyyy-MM-dd 06:00")
)
write-host "Provision-GovernanceTimerJobs.ps1 -Build $Build -Location $Location -AzureWebSite $AzureWebSite -JobCollection $JobCollection -StartDateAndTime $StartDateAndTime" -ForegroundColor Yellow
# DO NOT MODIFY BELOW
if($null -eq (Get-AzureAccount -EA 0)){
    Add-AzureAccount
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

function UploadJob($jobName,$jobFile, $jobType = "Triggered")
{
    Write-Host -ForegroundColor Yellow "Uploading $jobName"
    $site = Get-AzureWebsite -Name $AzureWebSite
    $job = $site | New-AzureWebsiteJob -JobName $jobName -JobType $jobType  -JobFile $jobFile
    if($jobType = "Triggered"){
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
            Remove-AzureSchedulerJob -Location $Location -JobCollectionName $jobCollectionName -JobName $jobName -Force | Out-Null
        }
        New-AzureSchedulerHttpJob -JobCollectionName $jobCollectionName -JobName $jobName -Method POST -URI "$($job.Url)\run" -Location $Location -StartTime $StartDateAndTime -Interval 1 -Frequency Day |Out-Null
        }
}
function DeployJob($packageName, $folder, $jobType="Triggered")
{
if($folder -eq $null)
{
    $folder= $packageName
}
# Check if the required release files are present
$files = Get-ChildItem "$basePath\OfficeDevPnP.PartnerPack.$folder\bin\$Build" -ErrorAction SilentlyContinue
if($files -ne $null -and $files.Length -gt 0)
{
    # Pack the files into a ZIP file
    $zipFile = Get-ChildItem "$basePath\OfficeDevPnP.PartnerPack.$folder\$packageName.zip" -ErrorAction SilentlyContinue
    if($zipFile -ne $null)
    {
        Remove-Item "$basePath\OfficeDevPnP.PartnerPack.$folder\$packageName.zip"
    }
    [IO.Compression.ZipFile]::CreateFromDirectory("$basePath\OfficeDevPnP.PartnerPack.$folder\bin\$Build","$basePath\OfficeDevPnP.PartnerPack.$folder\$packageName.zip");
    UploadJob "$packageName" "$basePath\OfficeDevPnP.PartnerPack.$folder\$packageName.zip" -jobType $jobType
} else {
    Write-Host -ForegroundColor Cyan "No build files available. Please configure and build the solution first."
}
}


$basePath = "$(convert-path ..)\OfficeDevPnP.PartnerPack.SiteProvisioning"


DeployJob "EnforceAdminsJob" "CheckAdminsJob"
DeployJob "ExternalUsersJob"
DeployJob "ScheduledJob" 
DeployJob "ContinuousJob" -jobType "Continuous"