Param(
    [Parameter(Mandatory = $true)] [String]$ProjectPath
#	[Parameter(Mandatory = $true)] [String]$PublishingSettingsPath
)

pushd $ProjectPath

# Get the path of MSBuild v. 14.0
$msbuildPath = Get-ItemProperty "hklm:\SOFTWARE\Microsoft\MSBuild\14.0"
$msbuildPath = $msbuildPath.MSBuildOverrideTasksPath.Substring(0, $msbuildPath.MSBuildOverrideTasksPath.IndexOf("\bin\") + 5)
$msbuildPath = $msbuildPath + "MSBuild.exe"

& $msbuildPath /p:Configuration=Release /p:VisualStudioVersion="14.0" 
# /p:PublishSettingsFile="$PublishingSettingsPath" /p:DeployOnBuild=true 
# /p:AllowUntrustedCertificate=true /p:_DestinationType=AzureWebSite

popd