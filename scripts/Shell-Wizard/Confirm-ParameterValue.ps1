[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Message explainining parameter to be populated.")]
    $prompt,
    [Parameter(Mandatory = $false, HelpMessage="Current Value. ")]
    $value,
     [Parameter(Mandatory = $false, HelpMessage="List of valid values. ")]
    $options = $null
)
if(-not [String]::IsNullOrEmpty($options))
{

    $value = $options | Out-GridView -Title $prompt -OutputMode Single
    return $value
}
if([String]::IsNullOrEmpty($value))
{
    while ([String]::IsNullOrEmpty($value)){
        $value =Read-Host "$prompt. Please, type in your desired value"

    }   
}else{
    $confirm = Read-Host "$prompt. Hit [Enter] for $value. Otherwise type in your desired value"
    if(-not [String]::IsNullOrEmpty($confirm))
    {
	    $value = $confirm 
    } 
}
return $value
