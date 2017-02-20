Param(
    [Parameter(Mandatory = $true)] [String]$ProjectPath,
	[Parameter(Mandatory = $true)] [String]$PublishingSettingsPath
)

pushd $ProjectPath

# Get the path of MSBuild v. 14.0
$msbuildPath = Get-ItemProperty "hklm:\SOFTWARE\Microsoft\MSBuild\14.0"
$msbuildPath = $msbuildPath.MSBuildOverrideTasksPath.Substring(0, $msbuildPath.MSBuildOverrideTasksPath.IndexOf("\bin\") + 5)
$msbuildPath = $msbuildPath + "MSBuild.exe"

# Create a temporary PnP path for the deployment files
$unique = (New-Guid).GetHashCode()
$tempPath = "$env:SystemDrive\PnP$unique"
mkdir $tempPath

# Run MSBuild to build and deploy the solution
& "$msbuildPath" /p:Configuration=Release /p:OutputPath="$tempPath" /p:VisualStudioVersion="14.0" /p:PublishSettingsFile="$PublishingSettingsPath" /p:DeployOnBuild=true

Remove-Item $tempPath -Recurse -Force
popd

