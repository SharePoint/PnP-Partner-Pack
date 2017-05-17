Param(
    [Parameter(Mandatory = $true)] [String]$ProjectPath,
	[Parameter(Mandatory = $true)] [String]$PublishingSettingsPath
)

pushd $ProjectPath

# Get the path of MSBuild v. 14.0.25420.1 or higher
$vs15 = Get-ItemProperty "hklm:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\SxS\VS7"
$msBuild15Path = $vs15.'15.0' + "\MSBuild\15.0\bin\msbuild.exe"
$vsVersion = "15.0"
if (Test-Path $msBuild15Path)
{
    $msbuildPath = $msBuild15Path
}

if ($msbuildPath -eq $null)
{
    $msbuildPath = Get-ItemProperty "hklm:\SOFTWARE\Microsoft\MSBuild\14.0"
    $msbuildPath = $msbuildPath.MSBuildOverrideTasksPath.Substring(0, $msbuildPath.MSBuildOverrideTasksPath.IndexOf("\bin\") + 5)
    $msbuildPath = $msbuildPath + "MSBuild.exe"
	$vsVersion = "14.0"
}

if ((Test-Path $msBuild15Path) -ne $true)
{
    Write-Host "Missing MSBuild"
}

# Create a temporary PnP path for the deployment files
$unique = (New-Guid).GetHashCode()
$tempPath = "$env:SystemDrive\PnP$unique"
mkdir $tempPath

# Run MSBuild to build and deploy the solution
& "$msbuildPath" /p:Configuration=Release /p:OutputPath="$tempPath" /p:VisualStudioVersion="$vsVersion" /p:PublishSettingsFile="$PublishingSettingsPath" /p:DeployOnBuild=true

Remove-Item $tempPath -Recurse -Force
popd


