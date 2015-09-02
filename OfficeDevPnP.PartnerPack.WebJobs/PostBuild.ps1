param($ProjectDir, $ConfigurationName, $TargetDir, $TargetFileName, $SolutionDir)

if($ConfigurationName -eq "Release")
{
	$DestinationZip = "$($SolutionDir)AdministrationJob.zip"

	If(Test-path $DestinationZip) 
	{
		Remove-item $DestinationZip
	}

	Add-Type -Assembly "System.IO.Compression.Filesystem"

	Write-Host "Zipping job files to $DestinationZip"

	[IO.Compression.ZipFile]::CreateFromDirectory($TargetDir, $DestinationZip) 
}